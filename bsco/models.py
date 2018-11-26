"""Models for the CRMM reader

The CRMM is the Meertens-created Corpus of ancient Dutch (14th century)

"""
import os, os.path
import sys
import io
import json
import requests

from util import ErrHandle

def get_exc_message():
    exc_type, exc_value = sys.exc_info()[:2]
    sMsg = "Handling {} exception with message '{}'".format(exc_type.__name__, exc_value)
    return sMsg

def downloadfile(url):
    """Downlaod a file from an URL to an (absolute) output file name"""

    # Default reply
    oBack = {}
    # Get the data from the URL
    try:
        r = requests.get(url)
    except:
        oBack['status'] = "error"
        oBack['code'] = "Cannot download from {}\nError: {}".format(url, get_exc_message())
        return oBack
    # Action depends on what we receive
    if r.status_code == 200:
        # Treat the reply as a complete string
        sText = r.text
        # Add the 'indices' separately
        oBack['text'] = sText
        oBack['status'] = 'ok'
    else:
        oBack['status'] = 'error'
        oBack['code'] = r.status_code
    # REturn what we have
    return oBack


class CrmmInfo():

    line = ""
    filenum = ""
    mrg_name = ""
    mrg_url = ""
    meta_url = ""
    location = ""

    def __init__(self, *args, **kwargs):
        """Create one item with information"""

        # Copy the information in the KWARGS
        for (k,v) in kwargs.items():
            setattr(self, k, v)

        self.oErr = ErrHandle()

        # perform other initialisations
        super(CrmmInfo, self).__init__()

    def get_json(self):
        """Return a JSON object"""

        obj = None
        try:
            obj = { 'line':  self.line,
                    'filenum': self.filenum, 
                    'mrg_name': self.mrg_name,
                    'mrg_url': self.mrg_url,
                    'meta_url': self.meta_url,
                    'location': self.location}
        except:
            sMsg = self.oErr.get_error_message()
        return obj
 
    def create_psd(self, targetdir):
        """Download the PSD from the link to the target directory"""

        oBack = {}

        try:

            # Figure out how the target file name should be
            fTargetPsd = os.path.abspath(os.path.join(targetdir, "crmm_{}.psd".format(self.filenum)))
            # Check if the file is already there
            if os.path.exists(fTargetPsd):
                self.oErr.Status("Skipping {}".format(fTargetPsd))
                oBack['status'] = 'ok'
            else:
                self.oErr.Status("Downloading meta of {}".format(fTargetPsd))
                # Get the text of the file
                oText = downloadfile(self.mrg_url)
                if oText['status'] == "ok":
                    # Get and save the text
                    sText = oText['text']
                    with io.open(fTargetPsd, "w", encoding="utf-8-sig") as f:
                        f.writelines(sText)
                    oBack['status'] = 'ok'
                else:
                    # There is an error
                    oBack = oText
        except:
            sMsg = self.oErr.get_error_message()
            oBack['status'] = "error"
            oBack['msg'] = sMsg
        return oBack

    def create_meta(self, targetdir):
        """Download the metadata from the link to the target directory"""

        oBack = {}

        try:
            # Figure out how the target file name should be
            fTargetPsd = os.path.abspath(os.path.join(targetdir, "crmm_{}.meta.xml".format(self.filenum)))
            # Check if the file is already there
            if os.path.exists(fTargetPsd):
                self.oErr.Status("Skipping {}".format(fTargetPsd))
                oBack['status'] = 'ok'
            else:
                # Get the text of the file
                self.oErr.Status("Downloading meta of {}".format(fTargetPsd))
                oText = downloadfile(self.meta_url)
                if oText['status'] == "ok":
                    # Get and save the text
                    sText = oText['text']
                    with io.open(fTargetPsd, "w", encoding="utf-8-sig") as f:
                        f.writelines(sText)
                    oBack['status'] = 'ok'
                else:
                    # There is an error
                    oBack = oText
        except:
            sMsg = self.oErr.get_error_message()
            oBack['status'] = "error"
            oBack['msg'] = sMsg
        return oBack
