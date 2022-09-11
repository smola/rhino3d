"""
Converts all .3ds files in a folder to a different Rhinoceros version.

It operates recursively, converting all subfolders.
"""

import rhinoscriptsyntax as rs
import scriptcontext as sc
import Rhino

import os
import os.path


def _log_error(msg):
    Rhino.RhinoApp.WriteLine("ERROR: " + msg)


def BatchConvert3dmVersion():
    # Switch to a new file, to ensure that the user has a chance to save any
    # unsaved progress.
    rs.Command("_-New _None", False)

    src_dir = rs.BrowseForFolder(message="Select folder to convert")
    if not src_dir or not os.path.isdir(src_dir):
        _log_error("Selected source folder is not valid")
        return

    dst_dir = rs.BrowseForFolder(message="Select destination folder")
    if not dst_dir or not os.path.isdir(dst_dir):
        _log_error("Selected destination folder is not valid")
        return

    # Pick a version between 1 and the current Rhino version.
    dst_version = rs.GetInteger("Select Rhino version", 5, 2, rs.ExeVersion())

    converted_count = 0
    error_count = 0
    # Walk through all directories under the source directory.
    for src_root, src_dirs, src_files in os.walk(src_dir):
        for src_file in src_files:
            # Ignore all files except .3dm
            if not src_file.endswith(".3dm"):
                continue

            # Get the absolute path of the original file and the converted file.
            src_path = os.path.join(src_root, src_file)
            rel_path = os.path.relpath(src_path, src_dir)
            dst_path = os.path.join(dst_dir, rel_path)

            # Create the destination directory if it does not exist.
            parent_dst_path = os.path.dirname(dst_path)
            if not os.path.exists(parent_dst_path):
                os.makedirs(parent_dst_path)

            # Skip if the destination file already exists.
            if os.path.exists(dst_path):
                continue

            Rhino.RhinoApp.WriteLine("Converting %s" % src_path)

            # Open the original file.
            #
            # Other appproaches:
            #   rs.Command ('_-Import "%s"' % (src_path), True)
            #   rs.Command ('_-Open "%s"' % (src_path), True)
            #
            if rs.ExeVersion() >= 7:
                doc = Rhino.RhinoDoc.OpenHeadless(src_path)
            else:
                doc = Rhino.RhinoDoc.Open(src_path)
            if not doc:
                _log_error("Could not open %s" % src_path)
                error_count += 1
                continue

            # SaveAs the new file with the new version.
            #
            # Other approaches:
            #   rs.Command('_-SaveAs _Version=%d "%s"' % (dst_version, dst_path), True)
            #
            result = doc.SaveAs(dst_path, dst_version)
            if not result:
                _log_error("Could not save %s" % dst_path)
                error_count += 1
                continue

            converted_count += 1

    rs.DocumentModified(False)
    rs.Command("_-New _None", False)
    rs.MessageBox(
        "Converted %d files.\nFound %d errors." % (converted_count, error_count)
    )


BatchConvert3dmVersion()
