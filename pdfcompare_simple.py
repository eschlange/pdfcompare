from pdfminer.pdfinterp import PDFResourceManager, process_pdf
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from cStringIO import StringIO
import re

def convert_pdf(path):

    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)

    fp = file(path, 'rb')
    process_pdf(rsrcmgr, device, fp)
    
    fp.close()
    device.close()

    str = retstr.getvalue()
    retstr.close()
    
    #filtered = filter(lambda x: not re.match(r'^\n\s*$', x), str)
    filtered = [line for line in str.split('\n') if line.strip() != '']
    
    print filtered

    return str
