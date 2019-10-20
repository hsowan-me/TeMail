#!/bin/bash

# generate setup.py
# pipenv run py2applet --make-setup TeMail.py

rm -rf build dist
pipenv run python setup.py py2app