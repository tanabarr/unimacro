__version__ = "$Revision: 356 $, $Date: 2011-02-08 11:14:12 +0100 (di, 08 feb 2011) $, $Author: quintijn $"
# (unimacro - natlink macro wrapper/extensions)
# (c) copyright 2003 Quintijn Hoogenboom (quintijn@users.sourceforge.net)
#                    Ben Staniford (ben_staniford@users.sourceforge.net)
#                    Bart Jan van Os (bjvo@users.sourceforge.net)
#
# This file is part of a SourceForge project called "unimacro" see
# http://unimacro.SourceForge.net).
#
# "unimacro" is free software; you can redistribute it and/or modify
# it under the terms of the GNU General Public License, see:
# http://www.gnu.org/licenses/gpl.txt
#
# "unimacro" is distributed in the hope that it will be useful, but
# WITHOUT ANY WARRANTY; See the GNU General Public License details.
#
# "unimacro" makes use of another SourceForge project "natlink",
# which has the following copyright notice:
#
# Python Macro Language for Dragon NaturallySpeaking
#   (c) Copyright 1999 by Joel Gould
#   Portions (c) Copyright 1999 by Dragon Systems, Inc.
#
# _number.py 
#  written by: Quintijn Hoogenboom (QH softwaretraining & advies)
#  August 2003
# 
"""smart number dictation

the number part of the grammar was initially provided by Joel Gould in
his grammar "winspch.py
it is changed a bit, but essentially all sorts of numbers can be
dictated with his grammar.

For real use copy the things from his grammar into another grammar. See
for example "_lines.py"

This grammar is initially off and should be so. When you want to try
this grammar, simply say "Switch On test number" ("Schakel test nummer
in").

QH211203: English numbers require more work: thirty three can be recognised as "33" or
as "30", "3".

QH050104: standardised things, and put functions in natlinkutilsbj, so that
other grammars can invoke the number grammar more easily.


"""
class NumberError(Exception): pass

import time, string, os, sys, types, re
import natlink
from actions import doKeystroke as keystroke

natut = __import__('natlinkutils')
natqh = __import__('natlinkutilsqh')
natbj = __import__('natlinkutilsbj')


# lists number1to99 and number1to9 etc. are taken from function getNumberList in natlinkutilsbj

##ancestor = natbj.TestGrammarBase
ancestor = natbj.IniGrammar
class ThisGrammar(ancestor):

    language = natqh.getLanguage()
    if language == "nld":
        name = "nummer"
        # het woord "Nummer" wordt niet makkelijk herkend, gebruik de
        # control toets om commando-herkenning te forceren, of wijzig het woord:
        gramSpec = """
<testnumber1> exported = (Getal [min] <number>)+;
<testnumber2> exported = Getal <number> tot <number>;
<number> = <integer> [(komma|punt) <integer>]; 
<pair> exported = geen betekenis <number>;  
# step 1 dutch:
"""+natbj.numberGrammar['nld']+"""
"""
##                       optioneel miljoen gedeelte (loopt minder soepel)
##                       | [<1to99> | <100to999>] miljoen [[en][<1to99> | <100to999>] duizend [[en](<1to99>|<100to999>)]]
    else:
        name = "number"
        # do not use "To" instead of "Through", because "To" sounds the same as "2"
        gramSpec = """
<testnumber1> exported = (Number [minus] <number>)+;
<testnumber2> exported = Number <number> Through <number>;
<pair> exported = combination <number>;  
# step 1 english: 
<number> = <integer> [(comma|point|dot) <integer>];   
"""+natbj.numberGrammar['enx']+"""
    """
  
    def initialize(self):
        if not self.language:
            print "no valid language in grammar "+__name__+" grammar not initialized"
            return
        self.load(self.gramSpec)
        # if switching on fillInstanceVariables also fill number1to9 and number1to99!
        self.switchOnOrOff() 

    def gotResultsInit(self,words,fullResults):
        # initialise the variables you want to collect the numbers in:
        self.number = self.through = ''
        self.minus = None
        print 'number: %s'% fullResults
        
    def gotResults_testnumber1(self,words,fullResults):
        # step 4: setting the number to wait for
        # because more numbers can be collected, the previous ones be collected first
        # if you expect only one number, this function can be skipped (but it is safe to call):
        self.collectNumber()
        print 'testnumber1: %s (%s)'% (words, self.number)
        if self.hasCommon(words, ['Number', 'Getal']):
            if self.number:
                # flush previous number
                if self.minus:
                    self.number = '-' + self.number
                print 'flush intermediate number: %s'% self.number
                self.outputNumber(self.number)
                self.number = ''
                self.minus = None
            self.minus = self.hasCommon(words, ['min', 'minus'])
            self.waitForNumber('number')
        else:
            raise NumberError, 'invalid user input in grammar %s: %s'%(__name__, words)

    def gotResults_pair(self,words,fullResults):
        self.waitForNumber('pair')

    def gotResults_testnumber2(self,words,fullResults):
        # step 4 also: if more numbers are expected,
        # you have to collect the previous number before asking for the new
        self.collectNumber()
        # can ask for 'number' or for 'through':
        if self.hasCommon(words, ['Number', 'Getal']):
            self.waitForNumber('number')
        elif self.hasCommon(words, ['Through', 'tot']):
            self.waitForNumber('through')
        else:
            raise NumberError, 'invalid user input in grammar %s: %s'%(__name__, words)


    def gotResults(self,words,fullResults):
        # step 5, in got results collect the number:
        self.collectNumber()

        if self.through:
            res = 'collected number: %s, through: %s'%(self.number, self.through)
            keystroke(res+'\n')
        elif self.number:
            if self.minus:
                self.number = '-' + self.number
            self.outputNumber(self.number)
        elif self.pair:
            keystroke("{f6}{f6}")
            num = self.pair*2 - 1
            keystroke("{tab %s}"% num)
            action("<<OpenPair>>")
            
                      
    def outputNumber(self, number):
        keystroke(number)
        prog = natqh.getProgName()
        if prog in ['iexplore']:
            keystroke('{tab}{extdown}{extleft}')
        elif prog in ['natspeak']:
            keystroke('{enter}')
        elif prog in ['excel']:
            keystroke('{tab}')

# standard stuff Joel (adapted for possible empty gramSpec, QH, unimacro)
thisGrammar = ThisGrammar()
if thisGrammar.gramSpec:
    thisGrammar.initialize()
else:
    thisGrammar = None

def unload():
    global thisGrammar
    if thisGrammar: thisGrammar.unload()
    thisGrammar = None 