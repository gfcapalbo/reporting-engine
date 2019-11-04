# -*- coding: utf-8 -*-
# © 2016 Therp BV <http://therp.nl>
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl.html).
from base64 import b64decode
from logging import getLogger
from PIL import Image
from StringIO import StringIO
from pyPdf import PdfFileWriter, PdfFileReader
from pyPdf.utils import PdfReadError
try:
    from PyPDF2 import PdfFileWriter, PdfFileReader  # pylint: disable=W0404
    from PyPDF2.utils import PdfReadError  # pylint: disable=W0404
except ImportError:
    pass
try:
    # we need this to be sure PIL has loaded PDF support
    from PIL import PdfImagePlugin  # noqa: F401
except ImportError:
    pass
try:
    from papersize import parse_couple, SIZES, UNITS
except ImportError:
    raise
from openerp import api, models, tools
logger = getLogger(__name__)


class Report(models.Model):
    _inherit = 'report'

    @api.multi
    def _resize_watermark(self, report, image):
        _format = report.paperformat_id.format.lower()
        # Returns pixels (width, height)
        im_width, im_height = image.size
        if _format == 'custom':
            # todo manage custom landscape case.
            pg_size = \
                report.paperformat_id.page_height, \
                report.paperformat_id.page_width
            ratio = pg_size[1] / pg_size[0]
        else:
            try:
                # returns point
                pg_size = parse_couple(SIZES[_format])
                # We convert to Inch, there are 72 points in an Inch
                pg_size = [float(x)/float(UNITS["in"]) for x in pg_size]
                # The ratio of every standard page-type is 0.7070
                # We resize our image on that ratio with minimal resize
                # ratio should allways be = 0.7070
                ratio = pg_size[0] / pg_size[1]
            except KeyError:
                logger.warning(
                    'Scaling the watermark failed.'
                    'Could not extract paper dimensions for %s' % (_format))
        # TODO, manage and verify landscape case
        new_im_size = (int(im_height * ratio), int(im_height))
        # the image will still bigger but in right proportion with minimal
        # loss of quality
        image_resized = image.resize(
            new_im_size, resample=Image.ANTIALIAS)
        # we calculate the DPI necessary to make image fit exactly
        # How many pixels in an Inch
        dpi = new_im_size[0] / pg_size[0]
        return image_resized, dpi

    @api.multi
    def _read_watermark(self, report, ids=None):
        if report.pdf_watermark:
            watermark = b64decode(report.pdf_watermark)
        else:
            watermark = tools.safe_eval(
                report.pdf_watermark_expression or 'None',
                dict(
                    env=self.env,
                    docs=self.env[report.model].browse(ids),
                )
            )
            if watermark:
                watermark = b64decode(watermark)
        if not watermark:
            return
        try:
            pdf_watermark = PdfFileReader(StringIO(watermark))
        except PdfReadError:
            image = Image.open(StringIO(watermark))
            pdf_buffer = StringIO()
            if image.mode != 'RGB':
                image = image.convert('RGB')
            dpi = 90
            if report.paperformat_id.format and \
                    report.paperformat_id.format.lower() in SIZES and \
                    report.pdf_watermark_scale:
                image, dpi = self._resize_watermark(report, image)
            # we save at 300
            image.save(
                pdf_buffer, 'pdf', subsampling=0, quality=95, resolution=dpi)
            pdf_watermark = PdfFileReader(pdf_buffer)
        return pdf_watermark

    @api.model
    def get_pdf(self, ids, report_name, html=None, data=None):
        report = self._get_report_from_name(report_name)
        result = super(Report, self).get_pdf(
            self.env[report.model].browse(ids),
            report_name,
            html=html,
            data=data,
        )
        pdf_watermark = self._read_watermark(report, ids)
        if not pdf_watermark:
            return result
        pdf = PdfFileWriter()
        for page in PdfFileReader(StringIO(result)).pages:
            watermark_page = pdf.addBlankPage(
                page.mediaBox.getWidth(), page.mediaBox.getHeight()
            )
            watermark_page.mergePage(pdf_watermark.getPage(0))
            watermark_page.mergePage(page)

        pdf_content = StringIO()
        pdf.write(pdf_content)
        return pdf_content.getvalue()
