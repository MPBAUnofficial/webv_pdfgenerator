from PyPDF2 import PdfFileWriter, PdfFileReader
import StringIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import json
import os
import sys


def fill_pdf(data, pdf_out='output.pdf'):
    """
    'empty_form.pdf' must be in the same directory of this script
    the script must be called with the target 'data.json' as the only argument
    """
    # load json data
    with open(data, 'r') as f:
        _data = json.loads(f.read())
    empty_form = PdfFileReader(open('empty_form.pdf', 'rb'))
    _output = PdfFileWriter()
    base_dir = os.path.dirname(os.path.abspath(data))

    packet = StringIO.StringIO()
    can = canvas.Canvas(packet, A4)

    # fill it with user data
    # Yeah, the coordinates are hard-coded... FML

    # last name
    can.drawString(60, 470.0, _data['profile']['last_name'])
    # first name
    can.drawString(60, 440.7, _data['profile']['first_name'])
    # birth date and birth place
    can.drawString(60, 411.4, '{0}, {1}'.format(
        _data['profile']['birth_place'], _data['profile']['birth_date']))
    # nationality
    can.drawString(60, 382.1, _data['profile']['nationality'])
    # address
    _a = _data['Home Address']
    address = '{0}{1}, {2}, {3}, {4}, {5}'.format(
        _a['Street 1'], ', ' + _a['Street 2'] if _a['Street 2'] else '',
        _a['City'], _a['State/Province'], _a['Country'], _a['Postal code'],
    )
    can.drawString(60, 352.8, address)
    # phone
    can.drawString(60, 323.5, _data['profile']['mobile_phone'])
    # mail
    can.drawString(60, 294.2, _data['profile']['email'])
    # school name
    can.drawString(60, 174.5, _data['School']['Name of the School'])
    # school's address
    can.drawString(60, 146, _data['School']["School's Address"])
    # reference professor
    can.drawString(60, 116.5, _data['School']['Reference Professor'])

    # draw the photo
    photo = _data['profile']['photo'].split('/')[-1]
    photo = os.path.join(base_dir, photo)
    can.drawImage(photo, 354.5, 526.7, 137, 152)
    can.save()

    # move to the beginning of the StringIO buffer
    packet.seek(0)
    new_pdf = PdfFileReader(packet)

    # the resulting pdf will be:
    # subscription form,
    # other (ordered) important forms,
    # all the remaining stuff
    first_forms = ('Code of Conduct.pdf', 'Parent Agreement.pdf',
                   'Assignment of Laptop.pdf', 'Media Consent Form.pdf')

    def append_pdf(pdf_in_path, pdf_out):
        pdf_in = PdfFileReader(file(pdf_in_path, 'rb'))
        [pdf_out.addPage(pdf_in.getPage(page_num))
         for page_num in range(pdf_in.numPages)]

    # generate the pdf
    page = empty_form.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    _output.addPage(page)

    signed_forms_dir = os.path.join(base_dir, 'signed-forms')
    # append all the important forms
    for form in first_forms:
        append_pdf(os.path.join(signed_forms_dir, form), _output)

    remaining_pdfs = [pdf for pdf in os.listdir(signed_forms_dir)
                      if os.path.isfile(os.path.join(signed_forms_dir, pdf))
                      and not pdf in first_forms]
    for pdf in remaining_pdfs:
        append_pdf(os.path.join(signed_forms_dir, pdf), _output)

    output_stream = file(pdf_out, 'wb')
    _output.write(output_stream)
    output_stream.close()

if __name__ == '__main__':
    if sys.argv <= 1:
        print 'Gimme the fucking data.json!'
        sys.exit(1)
    out = 'output.pdf' if len(sys.argv) < 3 else sys.argv[2]
    fill_pdf(sys.argv[1], out)
    print 'Done.'
