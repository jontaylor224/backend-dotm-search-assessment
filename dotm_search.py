#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "jontaylor224"

import argparse
import os
import zipfile


def get_docm_text(file_path):
    """ take path of docm file as argument, return the text"""
    document = zipfile.ZipFile(file_path)
    xml_content = document.read('word/document.xml')
    document.close()
    return xml_content


def main():
    matched_files = 0

    parser = argparse.ArgumentParser()
    parser.add_argument(
        '--dir',
        default=os.getcwd(),
        help='provide a file path to search (default: current directory)'
    )
    parser.add_argument(
        'search_term',
        help='provide a term to search '
    )
    args = parser.parse_args()

    print('Searching directory ' + args.dir + ' for text \'' + args.search_term + '\' ...')

    file_list = [os.path.join(args.dir, file) for file in os.listdir(
        args.dir) if file.endswith('.dotm')]

    for file in file_list:
        file_text = get_docm_text(file)
        if args.search_term in file_text:
            print 'Match found in file ' + file
            print '...' + file_text[file_text.index(args.search_term) -
                                    40:file_text.index(args.search_term)+len(args.search_term)+40] + '...'
            matched_files += 1
    print 'Total dotm files searched: ', len(file_list)
    print 'Total dotm files matched: ', matched_files


if __name__ == '__main__':
    main()
