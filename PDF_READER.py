from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from io import StringIO

def pdf_to_text(pdfname):
    try:
        from StringIO import StringIO
    except ImportError:
        from io import StringIO
    # PDFMiner boilerplate
    rsrcmgr = PDFResourceManager()
    sio = StringIO()
    device = TextConverter(rsrcmgr, sio,  laparams=LAParams())
    interpreter = PDFPageInterpreter(rsrcmgr, device)
      # get text from file
    fp = open(pdfname, 'rb')
    for page in PDFPage.get_pages(fp):
        interpreter.process_page(page)
    fp.close()
      # Get text from StringIO
    text = sio.getvalue()
      # close objects
    device.close()
    sio.close()

    print(text)

pdf_to_text("FV0415.pdf")
