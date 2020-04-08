import argparse

parser = argparse.ArgumentParser(description='Push local code to playground')
parser.add_argument('path_to_package', metavar='path_to_package', help='Path to the top of docassemble package (eg: /path/to/docassemble-packagename)')
parser.add_argument('--project', '-p', help='Docassemble playground project name')

args = parser.parse_args()

args
