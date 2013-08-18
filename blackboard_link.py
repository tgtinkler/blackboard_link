#!/usr/bin/python 
# Nomenclature
# PATH = bb/.foo.pdf
# DIRNAME = bb/
# BASE = .foo
# EXT = .pdf
# NAME = .foo.pdf

import os, sys
import subprocess

MS_OFFICE_EXTS = ['ppt', 'pptx', 'doc', 'docx']
ZIP_EXTS = ['zip']
SOURCE_DIR = 'bb'
IGNORE_FILE = '.ignored'

def split(path):
    dirname = os.path.dirname(path)
    if '.' in os.path.basename(path):
        base = '.'.join(os.path.basename(path).split('.')[:-1])
    else:
        base = os.path.basename(path)
    ext = path.split('.')[-1]
    return dirname, base, ext


def process(que):
    for command in que:
        subprocess.call(command, shell=True)


def link(path, label):
    """Path to source"""
    _, _, ext = split(path)
    if ext in MS_OFFICE_EXTS:
        target = convert_pdf(path)
    else:
        target = path
    os.symlink(target, label)


def convert_pdf(path):
    dirname, base, _ = split(path)
    convert_command = 'soffice --headless --convert-to pdf {}'.format(path)
    hidden_pdf = '{}/.{}.pdf'.format(dirname, base)
    move_command = 'mv "{}.pdf" {}'.format(base, hidden_pdf)
    que.append(convert_command)
    que.append(move_command)
    return hidden_pdf


def unpack_zips(dirname):
    """Unpack all the zips in a directory"""        

    for path in os.listdir(dirname):

        before = os.listdir(dirname)

        _, base, ext = split(path)

        if ext in ZIP_EXTS and '.'+base not in os.listdir(dirname):
            unzip_command = 'unzip "{}/{}" -d {}'.format(dirname, path, dirname)
            process([unzip_command])

            after = os.listdir(dirname)
            new = [name for name in after if name not in before]
            if len(new) == 1:
                move_command = 'mv "{}/{}" "{}/.{}"'.format(
                    dirname, new[0], dirname, base
                )
                process([move_command])


def walk_dir(root):
    for dirname, dirs, files in os.walk(root): 
        for filename in files:
            yield "{}/{}".format(dirname, filename)
            

def ignored():
        try:
            with open(IGNORE_FILE, 'r') as f:
                return [line.strip() for line in f.readlines()]
        except IOError:
            return []


def ignore(path):
    with open(IGNORE_FILE, 'a') as f:
        f.write(path+'\n')


def not_linked(SOURCE_DIR):
    """Return a list of files in dirname with no symlinks pointing to them"""

    targets = [
        split(os.readlink(path)) for path in walk_dir('.') 
        if os.path.islink(path)
    ]

    base_of_targets = [base for _, base, _ in targets]

    for path in walk_dir(SOURCE_DIR):
        _, base, ext = split(path)
        if all([
            base not in base_of_targets,
            base[0] <> '.',
            ext not in ZIP_EXTS,
            path not in ignored(),
        ]):
            yield path


if __name__ == '__main__':
    
    que = []
    
    unpack_zips(SOURCE_DIR)

    for path in not_linked(SOURCE_DIR):
        label = raw_input('{}: '.format(path))
        if label == 'i':
            ignore(path)
        elif label:
            link(path, label)

    process(que)
