#!/usr/bin/env python
from PyPDF2 import PdfFileWriter, PdfFileReader
import StringIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
import json
import os
from os import devnull
from os.path import abspath, join, dirname, isfile, isdir, exists, basename
import sys
import magic
from PIL import Image
import zipfile, rarfile
import argparse


class StudentRejectedException(Exception):
    """
    Why should you accept someone who's not able to create a freaking pdf?
    """
    pass


def append_pdf(pdf_in_path, pdf_out):
    """
    Merge two pdf files into one.
    """
    pdf_in = file_to_pdf(pdf_in_path)
    for page_num in range(pdf_in.numPages):
        pdf_out.addPage(pdf_in.getPage(page_num))


def archive_to_pdf(path, mimetype):
    """
    Helper method for file_to_pdf.
    Scan an archive and process all the files contained into it,
    avoiding (obviously) directories.
    """

    def zf_isdir(zf, name):
        """
        Return true if the given object inside the zipfile is a directory.
        (using os.listdir with zipfiles is not possible).
        """
        name = name.replace('\\', '/')  # fucking Windows
        return any(x.replace('\\', '/').startswith("%s/" % name.rstrip("/"))
                   for x in zf.namelist())

    files = []  # the files (not dirs) contained into the archive

    if mimetype == 'application/zip':
        with zipfile.ZipFile(path, 'r') as zf:
            for obj in zf.namelist():
                if not zf_isdir(zf, obj):
                    files.append(zf.read(obj, 'rb'))
    elif mimetype == 'application/x-rar':
        with rarfile.RarFile(path, 'r') as rf:
            for obj in rf.namelist():
                if not zf_isdir(rf, obj):
                    files.append(rf.read(obj, 'rb'))

    # create a void pdf
    output = PdfFileWriter()

    # insert images in a pdf and append it
    for f in files:
        pdf = file_to_pdf(f, is_buffer=True)

        # append it to the resulting pdf
        for page_num in range(pdf.numPages):
            output.addPage(pdf.getPage(page_num))

    # return the resulting pdf
    output_stream = StringIO.StringIO()
    output.write(output_stream)
    return PdfFileReader(output_stream, strict=False, warndest=os.devnull)


def file_to_pdf(file_in, is_buffer=False):
    """
    given a generic file (image, pdf, archive) return the PDF containing that
    file (adapted in order to fill properly).
    """

    if is_buffer:
        mimetype = magic.from_buffer(file_in, mime=True)
        file_in = StringIO.StringIO(file_in)
    else:  # then it's a path
        mimetype = magic.from_file(file_in, mime=True)

    if mimetype == 'application/pdf':
        if is_buffer:
            return PdfFileReader(file_in, strict=False, warndest=os.devnull)
        return PdfFileReader(open(file_in, 'rb'), strict=False, warndest=os.devnull)

    if mimetype and mimetype.split('/')[0] == 'image':
        image = Image.open(file_in)
        packet = StringIO.StringIO()
        can = canvas.Canvas(packet, A4)

        # if the image is too large, then resize it to fill in an A4 page
        if image.size[0] > A4[0] or image.size[1] > A4[1]:
            image.thumbnail(map(int, A4), Image.ANTIALIAS)

        # calculate the left and top margin to center the image in the page
        left, top = (A4[0] - image.size[0]) / 2, (A4[1] - image.size[1]) / 2

        # ...and draw it
        can.drawImage(ImageReader(image), left, top, preserveAspectRatio=True)
        can.save()
        packet.seek(0)
        return PdfFileReader(packet, strict=False, warndest=os.devnull)

    if mimetype in ('application/zip', 'application/x-rar'):
        return archive_to_pdf(file_in, mimetype)

    raise StudentRejectedException()


def fill_subscription_form(data, base_dir):
    """
    fill the subscription form (file 'empty_form.pdf') with user data.
    """
    empty_form = PdfFileReader(open('empty_form.pdf', 'rb'),
                               strict=False, warndest=os.devnull)
    packet = StringIO.StringIO()
    can = canvas.Canvas(packet, A4)

    # fill it with user data
    # Yeah, the coordinates are hard-coded... FML

    # last name
    can.drawString(60, 470.0, data['profile']['last_name'])
    # first name
    can.drawString(60, 440.7, data['profile']['first_name'])
    # birth date and birth place
    can.drawString(60, 411.4, '{0}, {1}'.format(
        data['profile']['birth_place'], data['profile']['birth_date']))
    # nationality
    can.drawString(60, 382.1, data['profile']['nationality'])
    # address
    _a = data['Home Address']
    address = '{0}{1}, {2}, {3}, {4}, {5}'.format(
        _a['Street 1'], ', ' + _a['Street 2'] if _a['Street 2'] else '',
        _a['City'], _a['State/Province'], _a['Country'], _a['Postal code'],
    )
    can.drawString(60, 352.8, address)
    # phone
    can.drawString(60, 323.5, data['profile']['mobile_phone'])
    # mail
    can.drawString(60, 294.2, data['profile']['email'])
    # school name
    can.drawString(60, 174.5, data['School']['Name of the School'])
    # school's address
    can.drawString(60, 146, data['School']["School's Address"])
    # reference professor
    can.drawString(60, 116.5, data['School']['Reference Professor'])

    def guess_photo_path():
        """
        Try to guess the path of the profile photo of a student where when it is
        not specified in data.json.
        """
        extensions = ['jpg', 'png', 'tiff', 'gif', 'jpeg']
        for ext in extensions:
            path = join(base_dir, 'profile.{0}'.format(ext))
            if exists(path):
                return path
        return None  # wtf, dude

    # try to guess the path of the profile picture
    if 'photo' in data['profile']:
        photo = data['profile']['photo'].split('/')[-1]
    else:
        photo = guess_photo_path()

    if photo is not None:
        photo = join(base_dir, photo)
        can.drawImage(photo, 354.5, 526.7, 137, 152)

    can.save()

    # move to the beginning of the StringIO buffer
    packet.seek(0)
    new_pdf = PdfFileReader(packet, strict=False, warndest=os.devnull)

    # generate the pdf
    page = empty_form.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    return page


def fill_pdf(data, output_dir=None, strict=False, verbose=False):
    """
    'empty_form.pdf' must be in the same directory of this script
    the script must be called with the target 'data.json' as the only argument
    If strict is True and the student did not submit ALL the important forms
    (the ones in the first_form tuple below) the pdf generation will stop.
    """
    # load json data
    with open(data, 'r') as f:
        _data = json.loads(f.read())

    _output = PdfFileWriter()
    base_dir = dirname(abspath(data))

    output_dir = abspath(output_dir) if output_dir is not None \
        else dirname(abspath(__file__))
    student_name = '{} {}'.format(_data['profile']['first_name'],
                                  _data['profile']['last_name'])
    output_file = join(output_dir, student_name + '.pdf')
    if output_file in os.listdir(output_dir):  # just a simple check
        output_file = '{} {}{}'.format(student_name,
                                       _data['profile']['email'], '.pdf')

    # the resulting pdf will be:
    # subscription form,
    # other (ordered) important forms,
    # all the remaining stuff
    first_forms = ('Code of Conduct', 'Parent Agreement',
                   'Assignment of Laptop', 'Media Consent Form')

    allowed_extensions = ('pdf', 'jpeg', 'jpg', 'png', 'tiff', 'gif')

    signed_forms_dir = join(base_dir, 'signed-forms')

    # set the first page as the (automagically filled) subscription form
    subscription_form = fill_subscription_form(_data, base_dir)
    _output.addPage(subscription_form)

    def _append_file(file_name, allowed_extensions):
        for ext in allowed_extensions:
            path = join(signed_forms_dir, '{0}.{1}'.format(form, ext))
            if exists(path):
                append_pdf(path, _output)
                return
        raise IndexError('File not found: {0}'.format(file_name))

    # append all the important forms
    for form in first_forms:
        try:
            _append_file(form, allowed_extensions)
        except IndexError:  # file not found
            if strict:
                print '\n{0} did not submit {1}. Skipping PDF generation.'\
                      .format(student_name, f)
                return

    remaining_files = [f for f in os.listdir(signed_forms_dir)
                       if isfile(join(signed_forms_dir, f))
                       and not '.'.join(f.split('.')[:-1]) in first_forms]

    for f in remaining_files:
        try:
            append_pdf(join(signed_forms_dir, f), _output)
        except StudentRejectedException:
            bastards.append(join(signed_forms_dir, f))
            if verbose:
                print "Unsupported file found: {0} , submitted by {1}"\
                    .format(join(signed_forms_dir, f), student_name)

    output_stream = file(output_file, 'wb')
    _output.write(output_stream)
    output_stream.close()


def fill_recursive(directory, output_dir, verbose=False, strict=False):
    """
    Given a directory, search recursively in that directory for 'data.json'
    files, and generate the related pdf.
    """
    for f in os.listdir(directory):
        _file = join(directory, f)

        if isfile(_file) and f == 'data.json':
            if verbose:
                print '\n+++ Elaborating file {}\n'.format(_file)
            try:
                fill_pdf(_file, output_dir, strict=strict, verbose=verbose)
            except KeyError, e:
                print 'Missing json field in {0}. PDF generation skipped. ' \
                      'Message: {1}'.format(_file, e.message)
                errors.append(dirname(_file))
            except Exception, e:
                print 'Something went wrong with {0}: {1}'\
                    .format(_file, e.message)
                errors.append(dirname(_file))

        if isdir(_file):
            if verbose:
                print 'Checking dir {}'.format(_file)
            fill_recursive(_file, output_dir, verbose=verbose, strict=strict)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='WebValley PDF generator')
    parser.add_argument('-r', '--recursive', dest='directory',
                        help='search recursively for data', nargs='?')
    parser.add_argument('-o', '--output-dir', help='set output directory',
                        nargs='?')
    parser.add_argument('data', help='data.json path', nargs='*')
    parser.add_argument('-v', '--verbose', help='verbose output',
                        action='store_true')
    parser.add_argument('-s', '--strict', action='store_true',
                        help='skip students who did not submit all the'
                             ' required forms')

    args = parser.parse_args()

    errors = []  # generic errors
    bastards = []  # people who submitted unsupported files (.doc, .odt, ...)

    if len(sys.argv) <= 1:
        parser.print_help()
        sys.exit(0)

    if args.output_dir:
        if not exists(abspath(args.output_dir)):
            print 'Error: destination dir {} does not exist.'\
                .format(abspath(args.output_dir))
            sys.exit(1)

    if args.directory is not None:
        if not args.directory or not exists(args.directory):
            print '>>> Gimme a VALID dir! <<<'
            sys.exit(1)
        fill_recursive(args.directory, args.output_dir, args.verbose,
                       args.strict)

    if args.data is not None:
        for json_path in args.data:
            # noinspection PyBroadException
            try:
                fill_pdf(json_path, args.output_dir or None,
                         args.strict, verbose=args.verbose)
            except Exception, e:
                if args.verbose:
                    print 'Error: {0}'.format(e.message)
                errors.append(dirname(json_path))

    if errors:  # generic problems
        path = join(args.output_dir or abspath(dirname(__file__)), 'errors.log')
        with open(path, 'w') as f:
            f.write('\n'.join(errors))
    if bastards:  # unsupported files found
        path = join(args.output_dir or abspath(dirname(__file__)),
                    'unsupported_files_found.log')
        with open(path, 'w') as f:
            f.write('\n'.join(bastards))
