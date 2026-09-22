# -*- coding: utf-8 -*-
"""Opens the website of VisionXt Service Italy"""

__title__ = 'VisionXt'
__author__ = 'Luca Rosati'
__context__ = 'zero-doc'

from pyrevit import script


# ponytail: plain pushbutton instead of a .urlbutton, whose bundle.yaml is fed
# to the python engine by pyRevit 6.0.0 master builds
script.open_url('https://visionxt.tech/')
