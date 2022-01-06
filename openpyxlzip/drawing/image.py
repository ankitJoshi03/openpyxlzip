# Copyright (c) 2010-2020 openpyxlzip

from io import BytesIO

try:
    from PIL import Image as PILImage
except ImportError:
    PILImage = False

from parse import parse

def _import_image(img):
    if not PILImage:
        raise ImportError('You must install Pillow to fetch image objects')

    if not isinstance(img, PILImage.Image):
        img = PILImage.open(img)

    return img


class Image(object):
    """Image in a spreadsheet"""

    _id = 1
    _path = "/xl/media/image{0}.{1}"
    anchor = "A1"

    def __init__(self, img):

        self.ref = img
        mark_to_close = isinstance(img, str)
        image = _import_image(img)
        self.width, self.height = image.size
        self.filename = None

        try:
            self.format = image.format.lower()
        except AttributeError:
            self.format = "png"
        if mark_to_close:
            # PIL instances created for metadata should be closed.
            image.close()


    def _data(self):
        """
        Return image data, convert to supported types if necessary
        """
        img = _import_image(self.ref)
        # don't convert these file formats
        if self.format in ['gif', 'jpeg', 'png']:
            img.fp.seek(0)
            fp = img.fp
        elif self.format in ['wmf']:
            img.fp.seek(0)
            fp = img.fp
        else:
            fp = BytesIO()
            img.save(fp, format="png")
            fp.seek(0)

        return fp.read()

    def extract_id(self):
        if self.filename is None:
            return self._id
        else:
            original_id, file_format = parse(self._path, "/" + self.filename)
            return original_id

    @property
    def path(self):
        return self._path.format(self._id, self.format)
