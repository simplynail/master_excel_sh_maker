# -*- coding: utf-8 -*-
"""
Created on Thu Feb 18 15:14:10 2016

@author: pawel.cwiek
"""


from distutils.core import setup
import py2exe
import sip

setup(windows=['ExcelMasterSpreadsheetMaker.py'], options={"py2exe":{"bundle_files":3,"includes":["sip","et_xmlfile"],"optimize": 2,"packages":["openpyxl"],"excludes":["matplotlib",'_gtkagg', '_tkagg','numpy','pandas', 'requests','doctest','pdb','unittest','difflib','nbformat','jinja2','PyQt5','nose','colorama','argparse','pydoc','asyncio','OpenGL','tornado','zmq']}})