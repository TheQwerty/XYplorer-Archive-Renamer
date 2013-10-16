import glob
import os
import shutil
import tempfile
import zipfile

from win32api import GetFileVersionInfo, LOWORD, HIWORD

def BatchRename():
    """ Rename all archives in the current directory. """

    # Create a temporary directory to store the extracted files.
    tempDir = tempfile.mkdtemp(prefix='XYplorer-')

    # Enumerate over all "*.zip" files.
    archives = glob.glob('*.zip')
    for thisArchive in archives:

        zippy = zipfile.ZipFile(thisArchive)
        try:
            # Extract the main executable.
            xyexe = zippy.extract('XYplorer.exe', path=tempDir)
        except:
            xyexe = None
        zippy.close()

        # Rename this archive from executable's version.
        RenameArchive(thisArchive, xyexe)

    # Remove temporary directory.
    try:
        shutil.rmtree(tempDir)
    except:
        print 'Could not remove tempDir: {0}'.format(tempDir)

def RenameArchive(thisArchive, xyexe):
    """ Rename an archive to the executable's version number. """

    if xyexe is None:
        raise TypeError
    if thisArchive is None:
        raise TypeError

    # Get file information.
    try:
        xyInfo = GetFileVersionInfo(xyexe, '\\')
    except:
        print 'No info...'
        return

    # Get various file version fields.
    xyFileVersion = getVersion(xyInfo, 'FileVersion')
    xyProductVersion = getVersion(xyInfo, 'ProductVersion')

    # Format FileVersion
    if xyFileVersion is not None:
        # Prefer FileVersion.
        result = versionStr(xyFileVersion)

    # Format ProductVersion
    if xyProductVersion is not None:
        # Use ProductVersion if there is no FileVersion.
        if result is None:
            result = versionStr(xyProductVersion)

        # Use both if they do not match.
        elif xyFileVersion != xyProductVersion:
            result = '{0}f_{1}p'.format(versionStr(xyFileVersion), versionStr(xyProductVersion))

    # Rename archive
    if result is not None:
        result = result.lower() + '.zip'
        base = os.path.basename(thisArchive).lower()

        print 'Renaming\t{0} >>\t{1}  '.format(thisArchive, result),

        if base == result:
            print 'Skipped'
        else:
            try:
                os.rename(thisArchive, result)
                print 'ok'
            except:
                print 'ERROR'

def getVersion(info, field):
    """ Convert file information into a meaningful tuple. """
    try:
        ms, ls = info[field + 'MS'], info[field + 'LS']
        return HIWORD(ms), LOWORD(ms), HIWORD(ls), LOWORD(ls)
    except:
        return None

def versionStr(v):
    """ Reformat version tuple as ##.##.####. """
    return "{0:02}.{1:02}.{3:04}".format(*v)

if __name__ == "__main__":
    BatchRename()