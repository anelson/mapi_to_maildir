#!/bin/sh
# Replaces the ';' with ':' in all files in the current branch of
# the filesystem.  Workaround to Windows' inability to store colons in
# file names
find . -name '*;*' -print | perl -w RestoreColons.pl 's/\;/\:/g'
