__version__ = "$Revision: 535 $, $Date: 2014-01-27 13:37:09 +0100 (ma, 27 jan 2014) $, $Author: quintijn $"
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
# natlinkutilsbj.py
#   This file contains utility classes and functions for grammar files.
#   Author: Bart Jan van Os; november 1999, adapted for
#   unimacro-project.
#   See the class BrowsableGrammar for documentation on the use of
#   the Grammar browser.
#   revised many times by Quintijn Hoogenboom
"""subclasses classes for natlink grammar files and utility functions

"""
import sys, os, string, cPickle, os.path, types, glob, time, re, shutil, copy
import natlink
import natlinkmain
import natlinkcorefunctions
#from natlinkutils import *  ## natut refers to natlinkutils
import gramparser # for translation with GramScannerReverse
from actions import doAction as action
from actions import doKeystroke as keystroke
from actions import Message
from actions import getAppPath, getAppName, getAppForEditExt
from actions import UnimacroBringUp

#consts = win32com.client.constants
import win32con
import win32api
import natlinkstatus
status = natlinkstatus.NatlinkStatus()
import spokenforms # for numbers spoken forms, IniGrammar (and also then DocstringGrammar)
# for translating the gramSpec in inigrammars:
reGrammarKeywords = re.compile(r"""
          (?<!{|<)        # not a { or a < before the string
          (\b[a-z]+\b|    # a whole word (only a-z)
          (?P<quote>['"])[a-z -]+(?P=quote))    # or 'more words' or "more words"
          (?!>|})         # not a > or a } after the string
          """, re.IGNORECASE+re.VERBOSE)
reSplitComments = re.compile(r"(#.*$)", re.MULTILINE)   # # through end of line

# boundaries for setList functions below:
grammarListLengthWarning = 55  # disabled 
grammarListLengthMaximum = 300

# difficult recognised counts:
# used in grammar tasks and firefox_browsing:
spokenFormCounts = {'nld': {'12': "twaalf"}}



import BrowseGrammar
# for IniGrammar:
import inivars
natqh = __import__('natlinkutilsqh')
natut = __import__('natlinkutils')
class UnimacroError(Exception): pass
# personal use: find out on what machine we are
def GetID():
    if os.environ.has_key('ID'):
        return string.lower(os.environ['ID'])
    else:
        return ''

# can be switched off:
automaticOpenFile = 1  # leave on!!
notifyIniErrors = 1 

# parameters for DisplayMessage
maxDisplayMessage = 75   # if messages longer, display in msgboxconfirm popup
DisplayMessageForbiddenCharacters = '\n\\~' # newline, backslash, tilde

# If you do not have DragonDictate running alongside, set the next
# constant to zero
DDIsRunning=0
LastMouseX=0
LastMouseY=0
if GetID()!='laptop':
    NaturalTextKey='{NumKey/}'
else:
    NaturalTextKey='{NumKey/}'
CorrectDlgKey='{NumKey-}'
DragonBarKey='{NumKey*}'   

# for search (in _general and _genericmovement)
lastSearchText = ''
lastSearchDirection = ''
beforeOrAfter = None  # valid values 'for', 'before', 'after'
app = None # special apps for search
appProgram = ''
comingFrom = ''

# special for UltraEdit32:
UEdit = None
uepath = r'C:\Program Files\UltraEdit\UEdit32.exe'
if os.path.isfile(uepath):
    UEdit = uepath
    print 'use UltraEdit32 for textual files open'# special for UltraEdit32:
import D_

def changeInList(L, old, new):
    """change in place"""
    if old in L:
        L.remove(old)
    L.append(new)

def getICAlphabet(language=None):
    """get the radio alphabeth, possibly with language changes
    when inserting "dummy" for old, a extra letter is added

    """
    A = D_.ICAlphabet[:]
    if language == 'nld':
        changeInList(A, 'a\\alpha', 'a\\alfa')
        changeInList(A, 'j\\juliet', 'j\\juliett')
        changeInList(A, 'x\\xray', 'x\\x-ray')
        #A = map(letterUppercase, A)
    return A

def getICAlphabetDict(language=None):
    """get the radio alphabeth, possibly with language changes
    when inserting "dummy" for old, a extra letter is added

    """
    A = getICAlphabet(language=language)
    if language == 'nld':
        changeInList(A, 'a\\alpha', 'a\\alfa')
        changeInList(A, 'j\\juliet', 'j\\juliett')
        changeInList(A, 'x\\xray', 'x\\x-ray')
        #A = map(letterUppercase, A)
    D = {}
    for k in A:
        wr, sp = k.split('\\')
        D[sp.strip()] = wr.strip()
    return D
    
def letterUppercase(l):
    """turns a\alpha into a\Alpha"""
    L = l.split('\\')
    return '%s\\%s'% (L[0], L[1].capitalize())

baseDirectory = natqh.getUserDirectory()

GrammarFileName=baseDirectory+'\\grammar.bin'
for somePath in sys.path:
    files = glob.glob(somePath + '\\pythonwin.exe')
    if files:
        PythonwinExe=files[0]
        break
else:
    print 'warning: no PythonwinExe found'
    
PythonwinPath, dummy = os.path.split(PythonwinExe)
PythonServerExe=os.path.join(PythonwinPath, 'pserver1')

# returns the union of two lists
def Union(L1,L2):
    if type(L1)==type(L2)==type([]):
        L=L1[:]
        for i in L2:
            if not i in L1: L.append(i)
        return L
    else:
        return []

# returns the intersection of two lists
def Intersect(L1,L2):
    if type(L1)==type(L2)==type([]):
        L=[]
        for i in L2:
            if i in L1: L.append(i)
        return L
    else:
        return []

def reverseDict(Dict):
    revDict={}
    for key in Dict.keys():
        revDict[Dict[key]]=key
    return revDict

def joinNestedStringLists(l):
    i=-1
    for sub in l:
        i=i+1
        if type(sub)==type([]): l[i]=joinNestedStringLists(sub)
    return string.join(l)

def getAllWords(fullResults):
    allWords=[]
    for word,rule in fullResults:
            allWords.append(word)
    return allWords

#Should return the operating system
def GetOS():
    if os.environ.has_key('OS'):
        return string.lower(os.environ['OS'])
    else:
        return ''

def IsWin98():
    #Don't know how. Better be on the safe side, return true if in doubt
    #Is used to skip actions that are problematic in Windows 98
    return  GetOS()!='windows_nt'

def ActivateDragonBarMenu():
    natut.playString("{NumKey-}",hook_f_systemkeys)        



def SetMic(state):
    # we avoid error messages, when turning the mic on while the device
    # is still busy.
    try:
        natlink.setMicState(state)
    except :
        pass

# mic switching
def SetDDMic(state):
    global DDIsRunning    
    if DDIsRunning:    
        try:
            if string.lower(state)=='on':
                if (GetOS()=='windows_nt'): #faster
                    natut.playString('{F12}')
                else:
                    natlink.execScript(' DdeExecute "Dragon Voicebar", "system", "[SetMicrophone 1]",1')
            elif string.lower(state)=='off':
                if (GetOS()=='windows_nt'): #faster
                    natut.playString('{F12}')
                else:                
                    natlink.execScript(' DdeExecute "Dragon Voicebar", "system", "[SetMicrophone 0]",1')
            else:
                print 'error: unknown DD mic state: ',state
        except:
            pass
    


def SwitchToDD():
    SetMic('off')    
    SetDDMic('on')

def SwitchToNat():
    global DDIsRunning    
    if DDIsRunning:        
        SetDDMic('off')
        Wait(1)
    SetMic('on')

# Natural text switching
def IsNaturalTextActive():
    (User,Dir)=natlink.getCurrentUser()
    try:
        OptFile=open(Dir+'\\'+'options.ini','r')
    except IOError:
        return 1
    OptLines = OptFile.readlines()
    OptFile.close
    for Line in OptLines:
        if Line[0:18]=='Global Dictation=0': return 0
    return 1

def ToggleNaturalText():
    natut.playString(NaturalTextKey,hook_f_systemkeys)
    
def SetNaturalText(State):
    State=string.lower(State)
    if GetDNSVersion()==5:
        stateNr=State=='on'
        natlink.execScript("SetNaturalText "+str(stateNr))
    else:
        if State=='on':
            if not IsNaturalTextActive(): ToggleNaturalText()
        elif State=='off':
            if IsNaturalTextActive(): ToggleNaturalText()

# mouse movement
def SaveMousePosition():
    global LastMouseX,LastMouseY
    (LastMouseX,LastMouseY)=natlink.getCursorPos()

def ReturnToLastMousePosition():
    natlink.playEvents([(wm_mousemove,LastMouseX,LastMouseY)])
    
def MoveMouseToRel(x,y):
    (cx,cy)=natlink.getCursorPos()
    natlink.playEvents([(wm_mousemove,x+cx,y+cy)])

def SetMousePosition(x,y):
    natlink.playEvents([(wm_mousemove,x,y)])    

#adapted from natlinkutils    
def ButtonClick(btnName='left',click='single'):
    x, y = natlink.getCursorPos()
    singleLookup = { 
        'left':  [(wm_lbuttondown,x,y),(wm_lbuttonup,x,y)],
        'right': [(wm_rbuttondown,x,y),(wm_rbuttonup,x,y)],
        'middle':[(wm_mbuttondown,x,y),(wm_mbuttonup,x,y)] }
    doubleLookup = {
        'left':  [(wm_lbuttondblclk,x,y),(wm_lbuttonup,x,y)],
        'right': [(wm_rbuttondblclk,x,y),(wm_rbuttonup,x,y)],
        'middle':[(wm_mbuttondblclk,x,y),(wm_mbuttonup,x,y)] }
    single = singleLookup[btnName]  # KeyError means invalid button name
    double = doubleLookup[btnName]

    if click == 'single': playEvents( single )
    elif click == 'double': playEvents( single + double )
    else: raise ValueError, "invalid click"

### lists for number grammar:
##language = natqh.getLanguage()
##print 'language: %s' % language
##if language == 'nld':
##    number1to99 = map(str, range(1,100))
##else:
##    number1to99 = map(str, range(1,21)) + ['30', '40', '50', '60', '70', '80', '90']
##number1to9 = map(str, range(1,10))

# app switching (is this used still?? getProg should be better, but is case insensitive QH2
def GetAppName():
#    if (GetOS()!='windows_nt'):
    App=getBaseName(natlink.getCurrentModule()[0])
#    else:
#        App=natlink.getCurrentModule()[0]
    return App


## movedgUp to natlinkutilsqh
## hope this is obsolete
##def AppSwapWith(App):
##        if (GetOS()!='windows_nt'):
##            app=string.lower(App)
##        else:
##            app=App
###    if not IsWin98():
##        try:
##            #seems to fail in my win98
##            #for some reason switching to TexPad fails if
##            # TeXPad is not opened with mouse from startmenu
###            natlink.execScript('AppSwapWith "'+app+'"')
##            natlink.execScript('AppBringUp "'+app+'"')            
##        except:
##            pass
##
# Here I define a common string search operation for different
# applications, using the common 'Find' dialog.
# When combined with the <dngwords> rule, you can
# provide a crude form of select( and say) in unsupported applications
def StartSearchFor(StringToSearch,Down=1,Full=0):
    moduleInfo = natlink.getCurrentModule()
    if matchWindow(moduleInfo, 'natspeak', '' ):
        natut.playString('{Ctrl+f}'+StringToSearch+'{Alt+f}{Esc}')
    elif matchWindow(moduleInfo, 'bibpad', '' ) or matchWindow(moduleInfo, 'textpad', '' ):
        if Down:
            if Full:
                natut.playString('{Ctrl+Home}')
            natut.playString('{F5}'+StringToSearch)
            natut.playString('{Alt+d}{Alt+f}{Esc}')
        else:
            if Full:
                natut.playString('{Ctrl+End}')
            natut.playString('{F5}'+StringToSearch)                
            natut.playString('{Alt+u}{Alt+f}{Esc}')

# Utility functions for global accessing of GrammarX objects.
# The CallAllGrammarObjects(funcName,args) method provides
# a means to send or receive signals/data to other grammar objects
# See also the GrammarX class
loadedGrammars = {}
grammarsChanged = 0

exclusiveGrammars = {}

def ClearGrammarsChangedFlag():
    """clears this flag, is used solely by grammar _control

    this flag was set when the grammar was registered or and registered

    """    
    global grammarsChanged
    grammarsChanged = 0

def RegisterGrammarObject(GrammarObject):
    """registers a grammar object in global variable

    also sets the flag "grammarsChanged"
    key in the dictionary loadedGrammars is the name,
    value is the instance object itself
    """    
    global loadedGrammars, grammarsChanged
##    print 'registering grammar object: %s: %s'% (GrammarObject.GetName(), GrammarObject)
    loadedGrammars[GrammarObject.GetName()] = GrammarObject
    grammarsChanged = 1

def UnRegisterGrammarObject(GrammarObject):
    """unregisters a grammar object from the global variable

    also sets the flag "grammarsChanged"
    delete the item in the dictionary "loadedGrammars"

    """    
    global loadedGrammars, grammarsChanged
    for k, v in loadedGrammars.items():
        if v is GrammarObject:
            del loadedGrammars[k]
            #print 'UNregistering grammar object: %s: %s'% (GrammarObject.GetName(), GrammarObject)
            grammarsChanged = 1
            break
    else:
        print 'cannot unregister grammar: %s'% GrammarObject

def CallAllGrammarObjects(funcName,args):
    """calls a function through all grammar objects

    funcName should be a string with the function name
    args should be a tuple of arguments, can be empty tuple ()
    exits silently if function doesn't exist in a grammar

    """    
    if args and len(args) == 1 and type(args[0]) == types.TupleType:
        args = args[0] # in order to be able to give arguments "loose" in the call instead of
                       # in a explicit tuple: CAGO(func, a, b, c) instead of
                       #                      CAGO(func, (a, b, c))
    for name, grammar in loadedGrammars.items():
        try:
            func = getattr(grammar, funcName)
            apply(func, args)
        except AttributeError:
            pass

def GetGrammarObject(grammarName):
    """return the grammar object, if in correct dict, by user name
    """
    if grammarName in loadedGrammars:
        return loadedGrammars[grammarName]

# Utility functions for displaying messages in the results box.
# The module _control registers these, and should be available.
# loadedMessageGrammar is a DictGrammar, which resides in _control. Is activated and
# deactivated only through DisplayMessage function.
#
# loadedControlGrammar is the grammar for control commands, which always run together with
# an exclusive mode, and can catch rejected recognitions when in exclusive mode.
# this control grammar never switches off.
loadedMessageGrammar = None
loadedControlGrammar = None
IsDisplayingMessage = 0

# if Display is asked for in gotBegin, it is waited for:
# displaying is controlled in _repeat:
pendingMessage = []


def SetPendingMessage(mess):
    global pendingMessage
    pendingMessage.append(mess)
def ClearPendingMessage():
    global pendingMessage
    pendingMessage = []
def GetPendingMessage():
    return pendingMessage

def SetDisplayingMessage():
    global IsDisplayingMessage
##    print 'setting display message'
    IsDisplayingMessage = 1

def ClearDisplayingMessage():
    global IsDisplayingMessage
##    print 'clearing display message'
    IsDisplayingMessage = 0

##def GlobalIsDisplayingMessage():
##    return IsDisplayingMessage
# instead of this function simply the variable natbj.IsDisplayingMessage can be checked

def RegisterMessageObject(GrammarObject):
    global loadedMessageGrammar
    loadedMessageGrammar=GrammarObject

def UnRegisterMessageObject(GrammarObject):
    global loadedMessageGrammar
    if (loadedMessageGrammar is GrammarObject):
        loadedMessageGrammar=None

def RegisterControlObject(GrammarObject):
    global loadedControlGrammar
    loadedControlGrammar = GrammarObject

def UnRegisterControlObject(GrammarObject):
    global loadedControlGrammar
    if (loadedControlGrammar is GrammarObject):
        loadedControlGrammar = None

def GlobalResetExclusiveMode():
    for k, v in exclusiveGrammars.items():
##        print 'resetting: %s'% k
        v[0].resetExclusiveMode()

##def Global(MessageText):
##    """displays message in the message window
##
##    works only if the grammar _message is active.
##
##    """    
##    global loadedMessageGrammar
##    if (loadedMessageGrammar):
##        try:
##            print 'global display message: %s'% MessageText
##            func = getattr(loadedMessageGrammar, 'DisplayMessage')
##            apply(func, [MessageText])
##        except AttributeError:
##            pass
##
##    
def  ControlResetExclusiveGrammar():
    """resets the exclusive flag of the _control grammar

    If the exclusive state of a/the last grammar is reset (becoming
    non-exclusive again), the grammar _control.
    This is done through this function

    """    
    
    global loadedContolGrammar
    if exclusiveGrammars:    # should not come here
        print 'no reset, still exclusiveGrammars present: %s'% exclusiveGrammars
        return
    if (loadedControlGrammar):
##        print 'resetting exclusive grammar control'
        loadedControlGrammar.setExclusive(0)

def ControlSetExclusiveGrammar():
    """sets the exclusive flag of the message grammar

    (the control grammar, including the message rule)

    This is to ensure the exclusive state together with
    another exclusive grammar. If in an exclusive grammar
    speech is not recognised, it is intercepted by the
    grammar _control, which invokes DisplayMessage.

    """    
    global loadedControlGrammar
    if not exclusiveGrammars:     # should not come here
        print 'no set of grammar, though exclusiveGrammars exist'
        return
    if loadedControlGrammar:
        if loadedControlGrammar.exclusive == 1:
##            print 'messageGrammar already exclusive'
            return
##        print 'setting exclusive grammar control'
        loadedControlGrammar.setExclusive(1)
  

#---------------------------------------------------------------------------
# GrammarX
# This is BJ's basic grammar class. It redefines some methods
# such that allways self.setExclusive is used to change the exclusive state.
# setExclusive is extended to keep track of the current exclusive state.
# It also registers itself and unregisters itself to accomadate the
# CallAllGrammarObjects function
# It also has the method DisplayMessage to display text in the results box, IF
#    the module _control (including a MessageDictationGrammar)
#    is available in the User directory of natpython

GrammarXAncestor=natut.GrammarBase # do not use self.__class__.bases[0]., see Python FAQ
class GrammarX(GrammarXAncestor):
    __inherited=GrammarXAncestor
    status = 'new'

    def __init__(self):
        self.__inherited.__init__(self)
        self.inGotBegin = 1 # initialise behave like being there
        self.mayBeSwitchedOn = 1
        self.isActive = 0
        self.language = natqh.getLanguage()
        self.version = natqh.getDNSVersion()
        self.exclusive = 0

    def getPrimaryAncestor(self):
        # the default primary ancestor is the first baseclass
        return self.__class__.__bases__[0] 

    def load(self,gramSpec,allResults=0,hypothesis=0, grammarName=None):
        if gramSpec:
            success = self.__inherited.load(self,gramSpec,allResults,hypothesis, grammarName=grammarName)
            if success: RegisterGrammarObject(self)
            return success

    def __str__(self):
        return '<grammarx: %s>'% self.GetName()

    def unload(self):
        UnRegisterGrammarObject(self)
        self.__inherited.unload(self)

    def beginCallback(self, moduleInfo):
        """overload from GrammarBase, does nothing if grammar is switched off

        """
        # intercept here if grammar is switched off, or inside displaying message
        if IsDisplayingMessage:
            return
        if self.mayBeSwitchedOn:
            # onOrOff = what the user wants
            # onOrOffState = the actual state.
            # if one of both is not False, gotBegin should cope with that
            # otherwise gotBegin can be skipped
            self.inGotBegin = 1
            self.callIfExists("gotBegin", (moduleInfo,) )
            self.inGotBegin = 0
            self.status = 'ready'

    def hypothesisCallback(self, words):
        # intercept here if grammar is switched off, or inside displaying message
        # when is this used?
        if self.mayBeSwitchedOn and not IsDisplayingMessage:
            self.callIfExists( "gotHypothesis", (words,) )

    def resultsCallback(self, wordsAndNums, resObj):
        self.__inherited.resultsCallback(self, wordsAndNums, resObj)

    # This is a utility function.  It calls a member function if and only
    # if that member function is defined.

    def getMode(self):
        """gets the current mode from natlinkmain

        is changed only in version 7!
        
        is only maintained when the mode is set with the function below.

        0 = normal, 1 = dictation, 2 = command, 3 = numbers, 4 = spell
        """
        return natlinkmain.DNSmode
    

    def setMode(self, n):
        """setting normal, dictation, command, numbers of spell mode

        remember the mode in instance variable currentMode.

        only for version 7        
        """
        print 'grammarX setMode: %s, version: %s'% (n, self.version)
        if self.version != 7:
            return

        old = self.getMode()
        
        if n == old:
            return
        if n not in range(5):
            print 'grammarX setMode: invalid mode: %s'% n
            return
        natlink.execScript('SetRecognitionMode %s'% n)
        # Bart Jan, is this too dirty?
        natlinkmain.DNSmode = n

    def getName(self):
        if "name" in dir(self):
            return self.name
        
        n = self.__module__
        if n[0] == "_":
            n = n[1:]
        self.name = n
        return n
    GetName = getName   # consistency with Bart Jan

    # These methods are adapted to keep track of the exclusive state.
    # They DO NOT intercept setExclusive! Bart Jan, I chose for calling once   QH11082003
    
    # but effectively call setExclusive twice to ensure original behaviour
    # under changing natlinkutils.
    def activate(self, ruleName, window=0, exclusive=None, noError=0):
        self.__inherited.activate(self, ruleName, window, noError=noError)
        if exclusive != None:
            self.setExclusive(exclusive)

    def deactivate(self, ruleName, noError=0):
        self.__inherited.deactivate(self, ruleName, noError)
        if not self.activeRules:
            self.isActive = 0
            

    def activateSet(self, ruleNames, window=0, exclusive=None):
        self.__inherited.activateSet(self, ruleNames, window)
        if not ruleNames:
            self.deactivateAll()
        self.isActive = 1
        if exclusive != None:
             self.setExclusive(exclusive)

    def activateAll(self, window=0, exclusive=None, exceptlist=None):
        self.__inherited.activateAll(self, window, exceptlist=exceptlist)
        self.isActive = 1
        if exclusive != None:
             self.setExclusive(exclusive)

    def deactivateAll(self):
        self.__inherited.deactivateAll(self)
        self.isActive = 0
        self.setExclusive(0)

    # used by GlobalResetExclusiveMode

    def resetExclusiveMode(self):
        # overload if normalSet is wanted:
        try:
            self.setExclusive(0)
        except natlink.NatError, t:
            print 'no resetExclusiveMode for %s: %s'% (self.name, t)

    def setExclusiveMode(self):
        """overload if special sets are wanted"""
        self.setExclusive(1)

    def setExclusive(self,exclusive):
        """let ControlGrammar follow exclusiveness

        state is maintained in self.exclusive
        """
        global exclusiveGrammars
        if exclusive == None:
            return
        if exclusive == self.exclusive:
            return
        self.__inherited.setExclusive(self,exclusive)
        self.exclusive = exclusive
        name = self.GetName()
##        print 'setting exclusive %s, self: %s'% (exclusive, self)
        if loadedControlGrammar and self is loadedControlGrammar:
            return
        if exclusive:
            if name in exclusiveGrammars:
                return
            moduleInfo = natlink.getCurrentModule()
            if self.activeRules:
                if len(self.activeRules) > 1:
                    rule = self.name
                else:
                    rule = self.activeRules[0]
            else:
                rule = '%s, no rules active'% self.name
            exclusiveGrammars[name] = (self, moduleInfo[2], rule)
            ControlSetExclusiveGrammar()
        elif exclusive == 0:
            if name in exclusiveGrammars:
                del exclusiveGrammars[name]
                if not exclusiveGrammars:
                    ControlResetExclusiveGrammar()
        else:
            print 'setExclusive, invalid value: %s (grammar: %s)'% (exclusive, name)

    def cancelMode(self):
        """overload for grammars that can go exclusive"""
        pass

    #Now we can get info on the exclusive state
    def getExclusiveInfo(self):
        name = self.GetName()
        if exclusiveGrammars and name in exclusiveGrammars:
            return exclusiveGrammars[name]

    def printExclusiveInfo(self):
        """global print info on exclusive grammar states"""
        if not exclusiveGrammars:
            print 'no grammars are exclusive'
        else:
            print 'exclusive are: %s'% exclusiveGrammars.keys()

    def switchOnOrOff(self, **kw):
        result = None
        if self.mayBeSwitchedOn == 'exclusive':
            print 'switch on exclusive: %s'% self.name
            if not (self.exclusive and self.isActive):
                result = self.switchOn(**kw)
                self.setExclusive(1)
        elif self.mayBeSwitchedOn:
            if (not self.isActive) or self.exclusive:
                result = self.switchOn(**kw)
        else:
            if self.isActive:
                result = self.switchOff()
        return result
            
    def switchOn(self, **kw):
        """switches grammar on, activates all rules

        must be overloaded if more specific behaviour is wished
        force = 1: set mayBeSwitchedOn
        """
        if 'force' in kw:
            if kw['force']:
                self.mayBeSwitchedOn = 1
            del kw['force']
        if self.mayBeSwitchedOn:
            self.activateAll(**kw)
            return 1
        print 'mayBeSwitchedOn False (%s), not switched on: %s'% (self.mayBeSwitchedOn, self.getName())

    def switchOff(self, force=None):
        """switches grammar off, deactivates all rules

        must be overloaded if more specific behaviour is wished
        """
        if force:
            self.mayBeSwitchedOn = 0
        self.deactivateAll()
        return 1

    def DisplayMessage(self,MessageText, pauseAfter=0, alert=None, alsoPrint=1):
        #if natqh.getDNSVersion >= 12:
        #    return # silently ignore this in Dragon 12
        if self.inGotBegin:
##            print 'message from gotBegin not allowed: %s'% MessageText
            SetPendingMessage(MessageText)
            return

        if not loadedMessageGrammar:
            print 'display message no message grammar: %s'% MessageText
            return
        if IsDisplayingMessage:
            print 'recursive call DisplayMessage, no display'
            return

        mayDisplay = 1
        for t in DisplayMessageForbiddenCharacters:
            if MessageText.find(t) >= 0:
                mayDisplay = 0
                break
        if len(MessageText) > maxDisplayMessage or not mayDisplay or alert:
            Message(MessageText, alert=alert)
            return
        if natqh.getDNSVersion() < 11:
            AddedMessage = '\\' + MessageText
        else:
            AddedMessage = '\\\\' + MessageText
        #print 'DisplayMessage: %s'% AddedMessage
        try:
            SetDisplayingMessage()
            if exclusiveGrammars:
                for g in exclusiveGrammars:
                    grammar = exclusiveGrammars[g][0]
    ##                print 'resetting exclusive: %s'% grammar
                    grammar.gramObj.setExclusive(0)
                if loadedControlGrammar:
                    loadedControlGrammar.gramObj.setExclusive(0)
##                print 'reset exclusive state in display message'

            hasWord = 0
            try:
                if (natlink.getWordInfo(AddedMessage)):
                    #print 'message already exists, text: %s' % MessageText
                    hasWord = 11
                else:
                    if natqh.getDNSVersion() < 8:
                        flags = natut.dgnwordflag_useradded | \
                                natut.dgnwordflag_no_space_next | \
                                natut.dgnwordflag_no_space_before
                        natlink.addWord(AddedMessage, flags)
                        #print 'added word: %s'% AddedMessage
                        hasWord = 1
                    else:
                        natlink.addWord(AddedMessage)
                        hasWord = 1
                        pass
##                        # flags for 8, 9...
##                        flagList = ['InternalUseOnly2', 'InternalUseOnly5',
##                                 'NoSpaceFollowingThisWord', 'NoSpacePreceedingThisWord',
##                                 'WordWasAddedByTheUser']
##                        flags = natqh.makeWordProperties(flagList)
##                    print 'flags: %x, intflags: %s'% (flags, int(flags))
            except natlink.InvalidWord:
                print 'error with addWord of: %s'% AddedMessage
 
            if hasWord:
##                print 'loadedMessageGrammar activate, with specific handle'
                hndle = natlink.getCurrentModule()[2]
                loadedMessageGrammar.activate(hndle)
                try:
##                    print 'send to messageDictGrammar: %s'% [AddedMessage]
                    natlink.recognitionMimic([AddedMessage])
                except (natlink.NatError, natlink.MimicFailed):
                    print 'error displaying message: %s'% AddedMessage

                # temporary changing the mode (only for version 7)                
    ##            if currentMode > 1:
    ##               natlink.execScript('SetRecognitionMode %s'% currentMode)
                if hasWord != 11:  # do not delete words that were already in voc
                    try:
                        natlink.deleteWord(AddedMessage)
                    except natlink.UnknownName:
                        pass  # for unknown reason a just added word sometimes cannot be deleted
                              # some other word in other capitalization?
                              #    (like \processing grammar and \Processing Grammar) ???
            else:
                # version 9 or error:
                loadedControlGrammar.setList('message',[AddedMessage])
                try:
##                    hndle = natlink.getCurrentModule()[2]
##                    loadedMessageGrammar.activate(hndle)
                    natlink.recognitionMimic([AddedMessage])
                    loadedControlGrammar.emptyList('message')
                except (natlink.NatError, natlink.MimicFailed):
                    print 'error displaying message (command): %s'% MessageText
        finally:
##            print 'loadedMessageGrammar deactivate'
            loadedMessageGrammar.deactivate()
            if exclusiveGrammars:
##                print 'set exclusive state again'
                for g in exclusiveGrammars:
                    grammar = exclusiveGrammars[g][0]
                    grammar.gramObj.setExclusive(1)
                if loadedControlGrammar:
                    loadedControlGrammar.gramObj.setExclusive(1)

##                loadedMessageGrammar.gramObj.setExclusive(1)
  
            ClearDisplayingMessage()
        if pauseAfter:
            print 'pause after DisplayMessage: %s'% pauseAfter
            time.sleep(pauseAfter)

    def message(self, mess):
        """only print to Messages window
        """
        print mess

    # progress report
    def IProgressReport(self,Case,TotalCases,NPeriods):
        Period=round(TotalCases/NPeriods)
        if (Case%Period==0) or (Case==TotalCases) or (Case==1):
            self.DisplayMessage('%3.0f%% Done' % (round(100.0*Case/TotalCases)), alsoPrint=0)

    def onTrayIcon(self, message):
        """needs this call when from actions the systemtray is called and clicked on

        see for example in _repeat"""
##        print 'caught onTrayIcon: %s'% repr(message)
        natlink.setTrayIcon()
        self.cancelMode()        

    def Wait(self, Time):
        """wait with the timer callback

        does not work when called from an action...
        
        """
        time.sleep(Time)
        return
    def wait(self, multiple=1):
        """wait a multiple of default times (default 0.1)
        """
        wTime = multiple * 0.1
        if wTime > 0:
            if wTime > 2:
                print 'waiting %s seconds'% wTime
            self.Wait(wTime)
        else:
            print 'no real wait, time is: %s'% wTime
##        # not succeed:::::
##        if Time > 5:
##            raise ValueError("self.Wait too long to respond (should be given in seconds): %s"% time)
##        self.didWait = 0
##        print 'starting with: %s'% Time
##        natlink.setTimerCallback(self.WaitOnce, int(Time*1000.0))
##        times = 0
##        while 0 <= self.didWait < 5 and times < 5:
##            times += 1
##            print 'sleeping, self.didWait: %s'% self.didWait
##        natlink.setTimerCallback(None)
##        print 'end of wait: %s'% getattr(self, 'interrupted', 0)
##            
##        
##    def WaitOnce(self):
##        """stop waiting again"""
##        print 'self.didWait: %s'% self.didWait
##        interrupted = getattr(self, 'interrupted', 0)
##        if self.didWait > 10 or interrupted:
##            print 'stopping: %s, %s'% (self.didWait, interrupted)
##            self.didWait = -1
##            natlink.setTimerCallback(None)
##        self.didWait += 1

#---------------------------------------------------------------------------
# BrowsableGrammar
#
# Class to support grammar browsing.
# It redefines some methods, such that copies of the used lists are
# being made.
# It has additional methods to support grammar browsing.
# It redefines resultsCallback to support an intercept mode:
# the grammar rules are recognized,
# but no callbacks are made, accept a call to gotInterceptedResults.
# That method displays the name of the grammar object and the rules used.
# => now in the browser.


BrowsableGrammarAncestor=GrammarX
class BrowsableGrammar(BrowsableGrammarAncestor):
    __inherited=BrowsableGrammarAncestor
    def __init__(self):
        self.__inherited.__init__(self)

    def load(self,gramSpec,allResults=0,hypothesis=0, grammarName=None):
        success = self.__inherited.load(self,gramSpec,allResults,hypothesis, grammarName=grammarName)
        if success:
            if type(gramSpec) == types.StringType:
                self.gramSpec=[gramSpec]
            else:
                self.gramSpec = gramSpec
            self.interceptMode = 0
            self.Lists={}
            for l in self.validLists:
                self.Lists[l]=[]
        return success

    def __str__(self):
        return '<browsablegr: %s>'% self.GetName()

    def switchOn(self, **kw):
        return self.__inherited.switchOn(self, **kw)

    def switchOff(self, **kw):
        return self.__inherited.switchOff(self, **kw)


    #The following two methods are redefined to keep a copy of the used lists
    def emptyList(self, listName):
        self.__inherited.emptyList(self,listName)
        self.Lists[listName]=[]

    def appendList(self, listName, words):
        """together with emptyList this is the setList command
        
        also check for long lists, > 100, warning, > 300 forbid
        QH, march 2013
        """
        if listName in self.Lists:
            lenone = len(self.Lists[listName])
        else:
            lenone = 0
        if type(words) == type(""):
            1
        else:
            lentwo = len(words)

        # may become a smaller list if grammarListLengthMaximum is exceeded:
        wordsToAppend = words
    
        if lenone + lentwo > grammarListLengthMaximum:
            if lenone:
                if lenone < grammarListLengthMaximum:
                    newAllowed = grammarListLengthMaximum - lenone
                    print 'appendList: can only append %s items (out of %s) to grammar list: "%s"'% (newAllowed, lentwo, listName)
                    wordsToAppend = words[:newAllowed]
                else:
                    print 'appendList: cannot anymore items (wanted %s) to grammar list: "%s", list already is at maximum size of %s'% \
                            (lentwo, listName, grammarListLengthMaximum)
                    return
            else:
                newAllowed = grammarListLengthMaximum
                print 'appendList/setList: can only set %s items (out of %s) to grammar list: "%s"'% (newAllowed, lentwo, listName)
                wordsToAppend = words[:newAllowed]
        elif lenone + lentwo > grammarListLengthWarning:
            pass  ## skip warning of large list:
            #if lenone:
            #    print 'appendList, warning list becoming large: appending %s items to grammar list: "%s", new length: %s'% (lentwo, listName, lenone+lentwo)
            #else:
            #    print 'appendList/setList, warning list becoming large: setting %s items to grammar list: "%s"'% (lentwo, listName)
            
            

        self.__inherited.appendList(self, listName, words)
        if type(words) == type(""):
            self.Lists[listName].append(words)
        else:
            for x in words:
                self.Lists[listName].append(x)

    # The next three methods implement the Intercept mode
    # for a 'Find Grammar' functionality.
    # Whenever InterceptMode is active, results are not transferred
    # to descendants, but are used to start the browser localized
    # to the recognized rules. Note that you do not need to set the
    # AllResults flag, because all grammars should be derived from this class.
    def setInterceptMode (self,Status):
        self.interceptMode = Status

    def gotInterceptedResults(self,words, fullResults):
        # Starts the browser with the intercepted results.
        # Note that it would be easy to make the browser only
        # display the CURRENT grammmar, but I prefer seeing all grammar.
        results = convertResults(fullResults)
        if string.join(words)!='Cancel This': #make sure one global grammar has this active!
            Start=(self.GetName(),results.keys())
            self.Browse(Start,0)
        CallAllGrammarObjects('setInterceptMode',[0])
        if string.join(words)=='Cancel This':
            self.DisplayMessage('Canceled', alsoPrint=0)

    def resultsCallback(self, wordsAndNums, resObj):
        # see natlinkutils for documentation. This method is changed
        # to implement interceptMode.
        # Uses the original behavior whenever NOT in intercept mode.
        # also intercept when in display mode:
        if IsDisplayingMessage:
##            print 'displaying message, ignore callback, grammar: %s'% self
            return
        elif not self.interceptMode:
            self.__inherited.resultsCallback(self, wordsAndNums, resObj)
        else:
           # do nothing more if the recog results were not for this grammar
           if type(wordsAndNums) != type([]):
               return None

           # we first convert the passed array of word/ruleNumbers into an
           # array of word/ruleNames and an array of only words
           words = []
           fullResults = []
           for x in wordsAndNums:
               words.append( x[0] )
               fullResults.append( ( x[0], self.ruleMap[x[1]] ) )
           #BJ ADD :
           self.gotInterceptedResults(words, fullResults)


    def GetDictionaries(self):
        # Overload to define Dictionaries; used in the grammar browser.
        # Suppose you have a rule with a variable part (List or enumeration
        # of alternatives), and the gotResults function uses a dictionary
        # that defines different actions to be taken, dependent on the variable
        # part: i.e. the variable part is used as the key value of the dictionary.
        # In such cases, it may be informative to display not only the
        # variable parts in the browser, but also the dictionary definitions.
        # You have to define the use of those dictionaries here.
        # The function should return one object, a map of all used dictionaries.
        # The key is either the rule name of the innermost rule
        # that uses the dictionary keys to define a (long) enumeration of
        # alternatives,
        #   <SomeRule> = [<RuleStart>] (""" + string.join(Dict.keys()," | ")+""");
        # OR the name of the list used within the rule,
        # associated with the dictionary:
        #   <ListRule> = [<RuleStart>] {Authors};
        #   plus
        #   self.setList('Authors',AuthorDict.keys())
        # The value for each rule/list name is the actual dictionary.
        # For example:
        #Dictionaries = {
        #    'SomeRule': Dict,        
        #    'Authors': AuthorDict,
        #     }
        #The default is empty.
        Dictionaries = {}
        return Dictionaries


    def InsertGrammarData(self,Grammars,All=0, Exclusive=0):
        #Parses the grammar specification and adds results to Grammars
        Name=self.GetName()
        if Name=='UnknownGrammar':
            Name=Name+str(len(Grammars.Included))
        Dicts=self.GetDictionaries()
        try:
            for n in Dicts.keys():
                if type(Dicts[n])!=type({}):
                    print Name+', '+n+': is not a dictionary'
        except:
            print Name+': Error in dictionary definitions'
        Elements=BrowseGrammar.ParseGrammarDefinitions(self.gramSpec,Name,self.Lists,
                                                       Dicts,self.activeRules,All,Exclusive,
                                                       self.exclusive)
        if Elements:
            # only if not empty, so leave out empty grammars (non-active if All = 0)
            Grammars.Append(Elements)

    def Browse(self,Start,All):
        self.DisplayMessage('Processing Grammar')
        BrowsePrepare(Start, All)
        self.DisplayMessage('Starting Browser')
        BrowseShow()

    def BrowsePrepare(self,Start,All, Exclusive=None):
        """only prepare the current state, should be followed by the BrowseShow
        
        """
        if Exclusive:
            print 'browse for exclusive grammars only'
            All = 0
            title = 'Active Rules in Exclusive Grammars'
        elif All:
            title = 'All Grammars'
        else:
            title = 'Active Grammars'
        Grammars=BrowseGrammar.GrammarElement()
        Grammars.Init(BrowseGrammar.RuleCode,title)
        CallAllGrammarObjects('InsertGrammarData',[Grammars,All,Exclusive])
        Data=(Grammars,Start,All,Exclusive)
        GrammarFile=open(GrammarFileName,'w')
        cPickle.dump(Data, GrammarFile)
        GrammarFile.close()

    def BrowseShow(self):
        """show the grammars as prepared in the function BrowsePrepare
        """
        pypath = '.'
        for x in sys.path:
            pypath = pypath + ';' + x
        os.environ['PYTHONPATH'] = pypath
        natqh.AppBringUp('Browser',Exec=PythonwinExe,Args='/app BrowseGrammarApp.py')


###
##GlobalGrammarBaseAncestor=BrowsableGrammar    
##class GlobalGrammarBase(GlobalGrammarBaseAncestor):
##    __inherited=GlobalGrammarBaseAncestor
##
##
##    def load(self,gramSpec,allResults=0,hypothesis=0):
##        print 'loading from GlobalGrammarBase: %s'% self.__module__
##        success = self.__inherited.load(self,gramSpec,allResults,hypothesis)
##        if success: self.setGlobalGrammarName()
##        return success
##
##
##    def setGlobalGrammarName(self):
##        try:
##            self.setList('GrammarName',self.GetName())
##        except:
##            pass
##
##    def gotResults_OpenThisGrammar(self,words,fullResults):
##        module=self.GetGrammarModuleName()
##        App=string.lower(getBaseName(getCurrentModule()[0]))
##        if not (App in ['pythonwin']):        
##            AppBringUp('Pythonw',Exec=PythonwinExe)
##        natut.playString('{Ctrl+l}'+module+'{Enter}')
##  
##LocalGrammarBaseAncestor=BrowsableGrammar    
##class LocalGrammarBase(BrowsableGrammar):
##    __inherited=LocalGrammarBaseAncestor
##    def gotResults_OpenThisGrammar(self,words,fullResults):
##        module=self.GetGrammarModuleName()
##        App=string.lower(getBaseName(getCurrentModule()[0]))
##        if not (App in ['pythonwin']):        
##            AppBringUp('Pythonw',Exec=PythonwinExe)
##        natut.playString('{Ctrl+l}'+module+'{Enter}')
##
IniGrammarAncestor=BrowsableGrammar    
class IniGrammar(IniGrammarAncestor):
    """grammar base which has methods for inifile stuff

    """
    __inherited=IniGrammarAncestor
    def __init__(self):
        self.__inherited.__init__(self)
        self.language = natqh.getLanguage()
        
        if not self.gramSpec:
            print 'Serious error: IniGrammar did not find gramSpec'
            return
        
        try:
            # done before, or coming from DocstringGrammar:
            self.gramSpec = copy.copy(self.originalGramSpec)
        except:
            # first time here, define self.originalGramSpec:
            if self.gramSpec:
                if isinstance(self.gramSpec, basestring):
                    self.originalGramSpec = self.gramSpec
                elif type(self.gramSpec) == types.ListType:
                    self.originalGramSpec = '\n'.join(self.gramSpec)
                else:
                    raise TypeError('IniGrammar "%s", gramSpec should be string or list, not: %s'% (self.name, type(self.gramSpec)))
        self.gramSpecTranslated = None  # a copy of the resulting translation if any (None otherwise)

        # here the grammar is not loaded yet, but the ini file is present
        # this could have been done in the user grammar already...
        if not 'ini' in dir(self):
            self.startInifile()
        self.name = self.checkName()
        self.debug = None # can be set in some grammar at initialize time
        #mod = sys.modules[self.__module__]
##            version = getattr(mod, '__version__', '---')
##            self.version = version.split(',')[0].strip("$Revision: ")
        self.version = None  #SVN change
        self.DNSVersion = natqh.getDNSVersion()
        self.spokenforms = spokenforms.SpokenForms(self.language, self.DNSVersion) # for spoken forms numbers!!

        #grammarName = grammarName or self.name

        translated = self.translateGrammar(self.originalGramSpec)  # pass original gramSpec...
        if translated:
            self.gramSpec = translated
            try:
                translationText = {'enx': ' (with synonyms)',
                                   'nld': ' (vertaald)'}[self.language]
            except KeyError:
                translationText = ' (translated)'
            self.nameForParser = self.name + translationText
        else:
            self.gramSpec = self.originalGramSpec
            self.nameForParser = self.name
            
    def __str__(self):
        return '<inigr: %s>'% self.GetName()

    def load(self,gramSpec,allResults=0,hypothesis=0, grammarName=None):
            grammarName = grammarName or self.nameForParser
            success = self.__inherited.load(self,gramSpec,allResults,hypothesis,grammarName)
            return success

    def checkName(self):
        """get possibly from inifile, if not present, start a inifile entry"""
        n = self.ini.get('grammar name', 'name')
        if n:
            return n
        if 'name' in dir(self):
            n = self.name
        else:
            n = self.__module__
        n = n.replace('_', ' ')
        n = n.strip()
        print 'setting grammar name to: %s'% n
        self.ini.set('grammar name', 'name', n)
        self.ini.write()
        return n

    def hasCommon(self, one, two, allResults=None, withIndex=None):
        """check if one and two have something in common

        one will often be the "words" list of the recognition

        two can be string or list, in both cases also in gramWords is searched too
        
      ((wish:  **as a special case, one can be a word of the original list, if used in a recursive callback! ))

        with allResults = 1, a list of possibly more hits is returned, always the words in the
        "two" list are returned.

        withIndex: return index of first hit in one (often the words)(not together with Allresults on)      
      
        """
        L = []
        i = -1
        if isinstance(two, basestring):
            two = [two]
        if isinstance(one, basestring):
            one = [one]
            
        for i, a in enumerate(one):
            aLower = a.lower()
            for b in two:
                bLower = b.lower()
                if aLower == bLower:
                    if allResults:
                        L.append(b)
                    elif withIndex:
                        return b, i
                    else:
                        return b
                elif self.gramWords and bLower in self.gramWords and aLower in self.gramWords[bLower]:
                    if allResults:
                        L.append(b)
                    elif withIndex:
                        return b, i
                    else:
                        return b
##        print 'result hasCommon: %s'% L
        if withIndex:
            return None, 0
        else:
            return L or None

    def translateGrammar(self, gramSpec):
        """translate all keywords with possible (multiple) words in your language


        translations (also optional or other keywords) can be entered in the section [grammar words]
        of the appropriate grammar file (say "edit name_of_grammar")

        separate wanted extra words (with the same meaning) by semicolons or |

        each keyword is searched for in the inifile                    
        
        """
        translateWords = self.getDictOfGrammarWordsTranslations() # from current inifile
        obsoleteTranslateWords = self.getDictOfObsoleteGrammarWords()
        self.ini.writeIfChanged() # marking unnecessery words
        oldTranslateKeys = translateWords.keys()
        newTranslateWords = {}
        self.gramWords = {} # keys: new grammar words, values mappings to the old grammar words

        if type(gramSpec) == types.ListType:
            self.oldGramSpec = '\n'.join(gramSpec)
            while type(self.oldGramSpec) == types.ListType:
                self.oldGramSpec = '\n'.join(self.oldGramSpec)
        else:
            # assume string:
            self.oldGramSpec = gramSpec
        if type(self.oldGramSpec) != types.StringType:
            print 'cannot translate gramSpec, wrong type: %s'% gramSpec
            return 
        localGramSpec = [self.oldGramSpec]
        gramparser.splitApartLines(localGramSpec)  # in the list splitting, gst gramSpecTranslated!

        # go and get the peek_ahead iterator of the GramScannerReverse:
        gsr = gramparser.GramScannerReverse(localGramSpec)
        gsgen = gsr.gramscannergen()
        it = gramparser.peek_ahead(gsgen)
        state = 'rule'  # 'rule' or 'expression'
        for wh, token, value in it:
            if state == 'rule':
                if token == '=':
                    # move to expression state
                    wordPrevParen = wordNextParen = 0
                    state = 'expr'
                gsr.appendToReturnList(wh, token, value)
                continue
            # now state == 'expr':
            if token == ';':
                # go back to a rule part:
                gsr.appendToReturnList(wh, token, value)
                state = 'rule'
                continue
            
            printedLines = 0
            # now inside a rule:
            if token in ['sqword', 'dqword', 'word']:
                if (not value.lower() in translateWords) and \
                        value.lower() in obsoleteTranslateWords:
                    if not printedLines:
                        print 'Make grammar obsolete word (translation/synonyms) current again: "%s": %s (grammar: "%s")'% (value.lower(),
                                                                                                                              obsoleteTranslateWords[value.lower()], self.name)

                    translateWords[value.lower()] = obsoleteTranslateWords[value.lower()]
                    del obsoleteTranslateWords[value.lower()]
                    self.ini.set('grammar words', value.lower(), translateWords[value.lower()])
                    self.ini.delete('grammar obsolete words', value.lower())
                    self.ini.write()
                    
                if value.lower() in translateWords:
                    if value.lower() in oldTranslateKeys:
                        oldTranslateKeys.remove(value.lower())
                    newValue = translateWords[value.lower()]
                    self.addToGramWords(value, newValue)
                    if len(newValue) == 0:
                        # no translation keep as it is
                        pass
                    elif len(newValue) == 1:
                        value = newValue[0]
                    else:
                        prevTok, nextTok = it.previous[1], it.preview[1]
                        parensNeeded = not (prevTok == '(' and nextTok in ['|', ')'] or
                                            prevTok == '[' and nextTok in ['|', ']'] or
                                            prevTok == '|' and nextTok in [']', ')', '|'])
                        #insert between parenthesis:
                        if parensNeeded:
                            gsr.appendToReturnList(wh, '(', None)
                        for w in newValue:
                            gsr.appendToReturnList('', 'word', w)
                            if w != newValue[-1]:
                                gsr.appendToReturnList('', '|', None)
                        if parensNeeded:        
                            gsr.appendToReturnList('', ')', None)
                        continue
                else:
                    # word with value in self.translateWords:
                    newTranslateWords.setdefault(value.lower(), set()).add(value)
            gsr.appendToReturnList(wh, token, value)
        self.correctIniFile(translateWords, newTranslateWords, obsoleteTranslateWords, oldTranslateKeys)
        self.stripDefaultTranlations(self.gramWords)
        if not self.gramWords:
            return # nothing is translated
        
        result = gsr.mergeReturnList()
        return result
            
    def addToGramWords(self, original, ListOfSynonyms):
        """build up the dict of gramWords synonyms/translations
        """
        orgLower = original.lower()
        L = self.gramWords.setdefault(orgLower, [])
        if ListOfSynonyms:
            for syn in ListOfSynonyms:
                if not syn.lower() in L:
                    L.append(syn.lower())
        else:
            L.append(orgLower)  # no translation, keep the original
                
                
    def stripDefaultTranlations(self, D):
        """strip from D (self.gramWords) the synonyms which are equal to the gramWord itself
        """
        for k, v in D.items():
            if len(v) == 1 and v[0] == k:
                del D[k]
        
            
    def correctIniFile(self, translateWords, newTranslateWords, obsoleteTranslateWords,
                       oldTranslateKeys):
        """correct if needed the sections [grammar words] and [grammar obsolete words]
        
        translateWords = dict of actual [grammar words] section, each value is a list of words
        newTranslateWords = dict of words that are NOT YET in the [grammar words] section, each value is a set of words
        obsoleteTranslateWords = dict of current obsolete grammar words, each value is a list of words
        oldTranslateKeys = keys from [grammar words] that are NOT used in the grammar, so have to be moved to the
                [grammar obsolete words] section
                
        """
        if oldTranslateKeys:
            # make words obsolete
            for k in oldTranslateKeys:
                if k == translateWords[k]:
                    self.ini.delete('grammar obsolete words', k)
                else:
                    # remember word translation for possible future use:
                    #print 'set obsolete: %s, %s'% (k, translateWords[k])
                    self.ini.set('grammar obsolete words', k, translateWords[k])
                self.ini.delete('grammar words', k)
            self.ini.write()
        notTransSection = 'grammar non translated words'
        
        if newTranslateWords:
            iniFileNonTranslatedWords = self.ini.getList(notTransSection, 'words')
            if iniFileNonTranslatedWords == newTranslateWords:
                pass
                #print 'correctIniFile, non translated words same as before'
            else:
                self.ini.set(notTransSection, 'words', newTranslateWords.keys())
                if self.language == 'enx':
                    self.ini.set(notTransSection, 'info1', 'These grammar words can be changed if you want synonyms for them.')
                elif self.language == 'nld':
                    self.ini.set(notTransSection, 'info1', 'De volgende grammatica woorden kunnen worden vertaald.')
                else:
                    self.ini.set(notTransSection, 'info1', 'These grammar words can be translated.')
                self.ini.set(notTransSection, 'info2', 'See http://unimacro.antenna.nl/features/translations for more info')
                self.ini.write()
        else:
            if self.ini.get(notTransSection):
                self.ini.delete(notTransSection)
                self.ini.write()
            
    ## for difficult recognised grammar words in counts:
    #def addSpokenFormNumbers(self, taskCounts):
    #    """if present, add spoken forms to the list of tasknumbers"""
    #    if self.language in spokenFormCounts:
    #        spoken = spokenFormCounts[self.language]
    #        for entry in spoken:
    #            if entry in taskCounts:
    #                taskCounts.append("%s\\%s"% (entry, spoken[entry]))
    #
    #def stripSpokenForm(self, text):
    #    return text.split("\\")[0]

    def getDictOfGrammarWordsTranslations(self):
        """return a dict of grammar word translations/synonyms
        
        keys: from inifile, grammar words section
        values: list of words in these values
        so: (in inifile)
        [grammar words]
        hello = hello
        there = there|already
        now = immediate;soon
        
        results in:
        {'there': ['there', 'already'], 'now': ['immediate', 'soon']}
        see unittestIniGrammar for more testing and examples.
        """
        DictFromIni = self.ini.toDict('grammar words')
        D = {}
        for k, v in DictFromIni.iteritems():
            k = k.lower()
            if type(v) == types.ListType:
                D[k] = v
            else:
                D[k]  = self.makeWordListFromGrammarWordsEntry(v)

            # ignore identical entries:
            vCheck = D[k]
            if len(vCheck) == 1 and vCheck[0] == k:
                self.ini.delete('grammar words', k)
                del D[k]
        return D

    def getDictOfObsoleteGrammarWords(self):
        """return a dict of grammar words which are marked obsolete in this grammar
        
        Just you can use them in case they are activated again, and they can be written back
        in case translation words (in the section [grammar words]) are not in the grammar any more)
        
        keys: from inifile, grammar words section
        values: list of words in these values
        so: (in inifile)
        [grammar words]
        there = there|already
        now = immediate;soon
        
        results in:
        {'there': ['there', 'already'], 'now': ['immediate', 'soon']}
        see unittestIniGrammar for more testing and examples.
        """
        DictFromIni = self.ini.toDict('grammar obsolete words')
        if not DictFromIni:
            return DictFromIni
        D = {}
        for k, v in DictFromIni.iteritems():
            k = k.lower()
            if type(v) == types.ListType:
                D[k] = v
            else:
                D[k]  = self.makeWordListFromGrammarWordsEntry(v)
            
            # ignore identical entries (capitalisation ignored): tot
            V = D[k]
            if len(V) == 1 and V[0] == k:
                self.ini.delete('grammar obsolete words', k)
                del D[k]
        if not D:
            self.ini.delete('grammar obsolete words')
        self.ini.writeIfChanged()
        return D

    def makeWordListFromGrammarWordsEntry(self, value):
        """get a entry from [grammar words] and split into a list
        
        input:  'hello; there'
        output: ['hello', 'there']

        """
        #key = self.cleanGrammarWord(key)
        #line = self.ini.get('grammar words', value)
        for splitchar in (',', ';', '|'):
            if value.find(splitchar) > 0:
                words = value.split(splitchar)
                words = map(self.cleanGrammarWord, words)
                break
        else:
            words = [self.cleanGrammarWord(value)]
            
        return filter(None, words)

                       

    def removeObsoleteGrammarWords(self):
        """remove from keywords list in ini file obsolete words

        if the grammar is changed, these words are not needed in the ini file any more.

        """
        allKeys = self.ini.get('grammar words')
        for k in allKeys:
            if k not in self.allGrammarKeywordsLower:
                v = self.ini.get('grammar words', k)
                self.ini.delete('grammar words', k)
                if k != v.lower():
                    self.ini.set('grammar obsolete words', k, v)
                self.iniChanged = 1

    def getTranslation(self, Dict):
        """get with self.language as key the text from dict, if invalid, return 'enx' text
        """
        if self.language in Dict:
            return Dict[self.language]
        else:
            return Dict['enx']


    def cleanGrammarWord(self, word):
        """strip and remove quotes

        this is from the input grammar into the inifile or self.gramWords, the words are inserted
        without surrounding quotes.

(test from testunimacro (qh))
>>> t = Test()
>>> print t.cleanGrammarWord('aap')
aap
>>> print t.cleanGrammarWord('"note mies"')
note mies
>>> print t.cleanGrammarWord("' noot    mies '")
noot mies

        """
        w = word.strip()
        w = word.strip('"\'')
        w = ' '.join(w.split())
        w = w.strip()
        return w
        
    def checkKeyword(self, t):
        """put quotes around if t has spaces, future: remove unwanted characters

        this is the way from the inifile into the grammar
        
        """
        if t.find(' ') > 0:
            return "'" + t + "'"
        elif t.find('-') > 0:
            return "'" + t + "'"  
        else:
            return t

    def switchOn(self, fillLists=1, **kw):
        """switches grammar on, activates all rules, fills lists

        this version assumes all lists are filled at switching on time.
        must be overloaded if more specific behaviour is wished
        """        
        # if you want rules dynamically activated in gotBegin,
        # the following line should be skipped in the overloaded function,
        # and the variable self.prevModInfo should be set to None
        if not self.__inherited.switchOn(self, **kw):
            print 'switching on "%s" failed'% self.name
            return
            

        # if you want to fill lists dynamically, you should skip
        # next line your overloaded function, and fill your lists
        # inside gotBegin, whenever self.prevModInfo is changed
        if not fillLists:
            # is used in grammar commands, see there
##            print 'skip filling lists: %s'% self.name
            return
        
        try:
            self.fillGrammarLists()
            if self.version:
                print 'IniGrammar switched on: %s (%s)'% (self.getName(), self.version)
            else:
                print 'IniGrammar switched on: %s'% self.getName()
                
        except:
            self.message('error switching on grammar %s, deactivate all rules\n\n (%s, %s)'% 
                               (self.getName(), sys.exc_info()[0], sys.exc_info()[1]))
            self.deactivateAll()            
    
    def showInifile(self, body=None, grammarLists=None, ini=None,
                    showGramspec=1, prefix=None, commandExplanation=None,
                    postfix=None, lineLen=60, sort=1):

        """gives grammar information and formats this in a file,
        that is opened after that.

        body: main text, can be left off (list or string)
        grammarLists: giving the lists from the inifile (if present),
                      that are used in the grammar
        ini: inivars instance, assumed to be self.ini
        showGramspec: 0 or 1, showing the complete grammar specification
        prefix: if given replaces the standard prefix
        commandExplanation: if given adds extra information (list or string)
        postfix: if given replaces the standard postfix

        beware:
        as I use language dependent grammars, I expect each grammar specification
        to start with 12 spaces!
        """
        ini = ini or self.ini

        # grammarLists can be given, or taken from the inifile:
        grammarLists = grammarLists or ini.get()
        activeLists = self.Lists.keys()
        moduleName = self.__module__
        # language must be used:
        language = natqh.getLanguage()
        assert type(grammarLists) == types.ListType
        grammarwordsLists = [l for l in grammarLists if l.startswith('grammar ')]
        grammarLists = [l for l in grammarLists if not l.startswith('grammar ')]
        grammarwordsLists.sort()
        grammarLists.sort()
        
        L  = []
        
        if prefix:
            L.append(prefix)
        else:
            try:
                formatLine = {'nld': 'Uitleg voor Unimacro (NatLink) grammatica "%s" (bestand: %s.py)'}[language]
            except:
                formatLine = 'Help for Unimacro (NatLink) grammar "%s" (file: %s.py)'
            L.append(formatLine % (self.name, moduleName))
        L.append('')

        if body:
            if type(body) == types.ListType:
                L.extend(body)
            else:
                L.append(body)
            L.append('')

        if grammarLists or activeLists:
            # get from ini, using again utilsqh:
            try:
                L.append({'nld': '\n--- lijsten:\n'}[language])
            except:
                L.append('\n--- lists:\n')
            grammarLists.sort()
            activeLists.sort()
            emptyLists = []
            for gList in activeLists:
                listValues = self.Lists[gList]
                if not listValues:
                    emptyLists.append(gList)
                    continue
                #if gList == 'degrees':
                #    print 'degrees: %s'% listValues
                #    print '----\ns2n: %s'% self.spokenforms.s2n
                #    print '----\nn2s: %s'% self.spokenforms.n2s
                onlyNumbers = self.spokenforms.sortedByNumbersValues(listValues, valueSpokenDict=1)
                if onlyNumbers:
                    #if gList == 'degrees':
                    #    print 'numbersDict: %s'% onlyNumbers
                    formattedTexts = inivars.formatReverseNumbersDict(onlyNumbers)
                    L.append("\n[%s]\n%s\n"% (gList, formattedTexts))
                elif '%sDict'%gList in dir(self):
                    if gList in grammarLists:
                        grammarLists.remove(gList)
                    dictName = '%sDict'%gList
                    items = sorted(getattr(self, dictName).keys())
                    L.append("\n[%s] (from %s)"% (gList, dictName))
                    L.append(formatListColumns(items))
                    L.append('\n')
                elif '%sList'%gList in dir(self):
                    if gList in grammarLists:
                        grammarLists.remove(gList)
                    listName = '%sList'%gList
                    items = sorted(getattr(self, listName))
                    listFromIni = sorted(ini.get(gList))
                    if listFromIni != items:
                        L.append("\n[%s] (from %s)"% (gList, listName))
                    else:
                        L.append("\n[%s]"% (gList))
                        items = listFromIni
                    L.append(formatListColumns(items))
                    itemsFromGrammar = sorted(self.Lists[gList])
                    if itemsFromGrammar != items:
                        L.append('\ncomplete list from grammar (different):')
                        L.append(formatListColumns(itemsFromGrammar))
                    L.append('\n')
                elif gList in grammarLists:
                    grammarLists.remove(gList)
                    L.append(ini.formatKeysOrderedFromSections([gList],
                                                   lineLen=lineLen, sort=sort))
                else:
                    L.append("\n[%s] (not found in inifile)"% gList)
                    items = sorted(self.Lists[gList])
                    L.append(formatListColumns(items))
                    L.append('')
            restSections = []
            for section in grammarLists:
                if section == 'general':
                    items = sorted(ini.get(section))
                    L.append("\n[%s] (options)"% section)
                    for i in items:
                        v = ini.get(section, i)
                        L.append('%s = %s'% (i, v))
                    #L.append(formatListColumns(items))
                    L.append('')
                else:
                    restSections.append(section)
            if restSections:
                L.append("\nOther sections:")
                L.append(formatListColumns(restSections))
                L.append('')
            if emptyLists:
                emptyLists.sort()
                L.append("\nEmpty lists:")
                L.append(formatListColumns(emptyLists))
                L.append('')
                
                
        if showGramspec:
            try:
                L.append({'nld': '\n--- grammatica:'}[language])
            except:
                L.append('\n--- grammar')
            if self.gramSpecTranslated:
                try:
                    L.append({'nld': ' --- vertaald', 'enx': ' --- with synonyms'}[language])
                except:
                    L.append(' --- translated')
                    
            L.append(' ---\n')
            if self.gramSpecTranslated:
                t = copy.copy(self.gramSpecTranslated)
            else:
                t = copy.copy(self.gramSpec)
                
            gramparser.splitApartLines(t)
                        
            #print 'gramSpec %s %s'% (type(t), repr(t))
            if t:
                if type(t) == types.ListType:
                    if len(t) == 1:
                        L.append(t[0])
                    else:
                        L.extend(t)
                else:
                    L.append(t)
            else:
                L.append('no grammar in self.gramSpec')

##        if grammarwordsLists:
##            continue
##            # get from ini, using again utilsqh:
##            # ignore these;;;;;;;;;;
##            try:
##                L.append({'nld': '\n--- grammatica lijsten (voor vertaling of synoniemen):\n'}[language])
##            except:
##                L.append('\n--- grammar lists (for translation or synonyms):\n')
##            L.append(ini.formatKeysOrderedFromSections(grammarwordsLists,
##                                                       lineLen=lineLen, sort=sort))
        active = self.activeRules
        valid = self.validRules
        active.sort()
        valid.sort()
        if valid and valid == active:
            pass
        elif active:
            L.append('active rules: %s'% ', '.join(self.activeRules))
        else:
            L.append('no rules are active')


        if commandExplanation:
            if type(commandExplanation) == types.ListType:
                L.extend(commandExplanation)
            else:
                L.append(commandExplanation)
            L.append('')
                    
     
        if postfix:
            L.append(postfix)
        else:
            try:
                formatLine = {'nld': '\n--- gebruiker: %s, %s'}[language]
            except:
                formatLine = '\n--- user:- %s, %s'
            L.append(formatLine %(natlink.getCurrentUser()[0], time.asctime(time.localtime(time.time()))))

        try:
            extension = {'nld': '_uitleg.txt'}[language]
        except:
            extension = '_help.txt'

        
        whatFile = os.path.join(os.environ['TEMP'], language + '__' + moduleName.replace(' ', '_') + extension)
        if os.path.isfile(whatFile):
            #print 'already exist, try to remove: %s'% whatFile
            try:
                os.remove(whatFile)
            except IOError:
                print 'Cannot remove previous help (show) file: "%s:"\nProbably this file is still open in Notepad\nPlease close and call your command again'% whatFile
                return
        if os.path.isfile(whatFile):
            print 'Cannot remove previous help (show) file: "%s:"\nProbably this file is still open in Notepad\nPlease close and call your command again'% whatFile
            return
        
        #print 'writing to and open:\n\t"%s"'% whatFile
        fSock = open(whatFile, 'w')
        
        fSock.write('\n'.join(L))
        self.openFileDefault(whatFile, mode="edit")

    def editInifile(self):
        """opens the inifile that controls the behaviour of
        the grammar
        
        if inifile is not given, the standard name is expected
        """
        inifile = self.inifile
        self.iniFileDate = natqh.getFileDate(inifile)
        self.checkForChanges = 1
        self.openFileDefault(inifile, mode="edit")

    def fillGrammarLists(self, listOfLists=None, ignoreFromIni='general',
                         ignoreFromGrammar=None):
        """fills the lists of the grammar with data from inifile

        If listOfLists is not provided, all sections from the inifile
        are tried.

        Lists "number1to99" and "number1to9" are taken from this file, unless
        they are provided in inifile
        As well as "n1-9" etc
        
        Numbers are taken from spokenforms, as well as 'character'.

        If a list of names is given in "ignoreFromIni", or this variable is
        a valid section name, this/these name(s) are ignored in this function

        Also if listName is found in self.iniIgnoreGrammarLists, it is skipped,
        but with a warning

        If "ignoreFromGrammar" is or contains valid list names in the grammar,
        these are ignored when filling the lists.

        At the end of the function it is checked if there are no
        remaining lists from the grammar that should be filled.

        """
        if not self.ini:
            raise UnimacroError('no ini file active for grammar: %s'% self.GetName())
        ini = self.ini
        fromGrammar = self.validLists[:]
        allListsFromIni = ini.get()
        nonemptyListsFromIni = [l for l in allListsFromIni if ini.get(l)] # empty sections ignored
        if 'iniIgnoreGrammarLists' in dir(self):
            self.removeFromList(fromGrammar, self.iniIgnoreGrammarLists)
        if listOfLists:
            for l in listOfLists:
                if l not in fromGrammar:
                    self.error('fillGrammarLists, list name not in grammar: %s'% l)
                    continue
                if l not in listsFromIni:
                    self.error('fillGrammarLists, list name not in ini file: %s'% l)
                    continue
                if self.fillList(l):
                    fromGrammar.remove(l)
                
        else:
            for l in nonemptyListsFromIni:
                if l not in fromGrammar:
                    continue
                if self.fillList(l):
                    fromGrammar.remove(l)

        if fromGrammar:
            changes = 0
            for l in fromGrammar[:]:
                if self.fillList(l):
                    fromGrammar.remove(l)
                elif l in allListsFromIni: 
                    print '%s: warning, empty section for list "%s" in inifile'% (self.name, l)
                else:
                    changes = 1
                    self.emptyList(l)
                    self.ini.set(l)
            if changes:
                self.ini.write()
                commandText = {'nld': 'bewerk'}
                try:
                    commandWord = commandText[self.language]
                except KeyError:
                    commandWord = "edit"

                self.message('fillGrammarLists in grammar "%s"\n\nNot all lists filled: %s\n\nPlease fill in in inifile by calling the command "%s %s"'%
                                    (self.name, fromGrammar, commandWord, self.name))
                self.checkForChanges = 1
                self.iniFileDate = natqh.getFileDate(self.inifile)
                
                #self.openFileDefault(self.inifile)

    def setNumbersList(self, name, numbers):
        """try to get spoken forms for the numbers
        """
        mixed = self.spokenforms.getMixedList(numbers)
        #print 'mixed list (setNumbersList) of %s: %s'% (name, mixed)
        self.setList(name, mixed)
        return mixed # so true if list is there...

    def setCharactersList(self, name, chars=None):
        """try to get spoken forms for the characters, a-z if chars is None
        """
        if chars is None:
            chars = string.ascii_lowercase
            
        spokenList = self.spokenforms.getMixedCharactersList(chars)
        #print 'characters list "%s": %s'% (name, spokenList)
        self.setList(name, spokenList)
        return spokenList
            
    def getCharacterFromSpoken(self, word, originalList=None):
        """returns a character, corresponding with word
        
        if originalList is given ('abcd...') validate the result with this list
        """
        if isinstance(word, basestring):
            return self.spokenforms.getCharFromSpoken(word, originalList)
        else:
            raise ValueError('getCharFromSpoken: input must be a string (normally one of the "words", not: %s'% word)

    def getNumberFromSpoken(self, word, originalList=None, asStr=None):
        """returns as the number, corresponding with word
        
        if originalList is given (either list of int or str), validate the result with this list
        if asStr = True, return as string, otherwise return as int
        """
        if isinstance(word, basestring):
            return self.spokenforms.getNumberFromSpoken(word, originalList, asStr)
        else:
            raise ValueError('getNumbersFromSpoken: input must be a string (normally one of the "words", not: %s'% word)
            
    def getNumbersFromSpoken(self, words, originalList=None, asStr=None):
        """returns as list the numbers from wordOrWords
        
        if originalList is given (either list of int or str), validate the result with this list
        if asStr = True, return as string, otherwise return as int
        """
        if type(words) == types.ListType:
            L = [self.spokenforms.getNumberFromSpoken(w, originalList, asStr) for w in words]
            L = filter(None, L)
            return L
        else:
            raise ValueError('getNumbersFromSpoken: input must be a list (normally "words", not: %s'% words)

    def fillList(self, listName):
        """fill a list in the grammar from the data of the inifile

        """
        n = listName
        if n[:6] == 'number' or \
                  (n[0] == 'n' and '-' in n):
            L = self.spokenforms.getNumberList(n)
            if L:
                #print 'numbers (of %s): %s'% (n, L)
                return self.setNumbersList(n, L)
        elif n == 'character':
            print 'character list'
            return self.setCharactersList('character')
        else:
            l = self.ini.get(n)
            if l:
                #if self.debug:
                #print '%s: filling list %s with %s items'% (self.name, listName, len(l))
                self.setList(n, l)
                return 1
            else:
                #if self.debug:
                #print '%s: not filling list %s, no items found'% (self.name, listName)
                self.emptyList(n)

    def fillInstanceVariables(self):
        """fills instance variables with data from inifile

        default: nothing happens, must be supplied by the calling grammar    

        """
        pass

    def startInifile(self):
        """loads the inifile for the grammar given.

        Is called at initialisation of this instance.

        In this function the following variables are set,
        checkForChanges, starts normally at 0
            if you perform the "editGrammar" function, this variable is set to 1,
            and the gotBegin function only goes into the function "checkInifile",
            if this variable is set.

        if the commandDir contains a folder "private", AND the specified file
        is in this folder, this inifile is taken. meant only for developers.

        After this function also are set:
        ini, the inivars instance that controls the ini variables
        inifile, the full path of the inifile that is opened
        inifileDate, the date/time the inifile was last modified        

        """
        # get default inifile name:
        self.checkForChanges = 1
        self.openedInifile = 0
        self.ignore = None
        modName = self.__module__
        baseDir = natqh.getUserDirectory()
        userDir = natqh.getUnimacroUserDirectory()
        originalUnimacroDir = natqh.getOriginalUnimacroDirectory()
        
        commandDir = os.path.join(userDir,
                                        self.language +"_inifiles")
        inifile = os.path.join(commandDir, modName + '.ini')
        #print 'starting inifile for %s'% modName
        if not os.path.isfile(inifile):
            print 'Cannot find inifile: %s'% inifile
            self.lookForExampleInifile(commandDir, modName + '.ini')
            if not os.path.isfile(inifile):
                print 'cannot find an example inifile for %s'% modName
                self.inifile = inifile
                self.TryToMakeDefaultInifile(commandDir,modName + '.ini', self.language)
                if not  os.path.isfile(inifile):
                    self.message('cannot make a default inifile for: %s (%s)'%
                                        (modName, inifile))
                    self.inifile = None
                    return
            #self.openFileDefault(inifile)
            commandText = {'nld': 'bewerk'}
            try:
                commandWord = commandText[self.language]
            except KeyError:
                commandWord = "edit"
            self.name = self.getName()
            self.message('===Created new inifile for grammar "%s"\n===Please edit this file by calling the command "%s %s"'%
                         (modName, commandWord, self.name))
        self.inifile = inifile
        #self.ini = inivars.IniVars(self.inifile, repairErrors=1)

        self.iniFileDate = natqh.getFileDate(self.inifile)
        try:    
            self.ini = inivars.IniVars(self.inifile, repairErrors=1)
        except inivars.IniError:
            mess = str(sys.exc_info()[1])
            self.openFileDefault(self.inifile)
            print '****IniError while initialising: %s' % self.name
            if notifyIniErrors:
                self.message(mess)
            print '****error message: '+mess

            self.ini = None
            return


        # starting variable:
        self.prevModInfo = None        
        self.checkForChanges = 0
        
        # control on or off:
        initialOn = self.ini.get('general', 'initial on', '1')
        user = natqh.getUser()
        initialOnUser = self.ini.get('general', 'initial on %s'% user)
        initialOn = initialOnUser or initialOn
        if initialOn.lower() == 'exclusive':
##            print 'ask for exclusive: %s'% self.name
            self.mayBeSwitchedOn = initialOn.lower()
        else:
            self.mayBeSwitchedOn = (initialOn.lower()  in ['1', 'on'])

        self.prevModInfo = None
        try:
            self.fillInstanceVariables()
        except inivars.IniError:
            mess = str(sys.exc_info()[1])
            self.openFileDefault(self.inifile)
            if notifyIniErrors:
                self.message(mess)
            print '***IniError when filling instance variables: %s' % self.name
            print '***error message: '+mess
            self.ini = None
            return

    def lookForExampleInifile(self, commandDir, fileName):
        """must be tested, look for a valid inifile in one of the sample dirs
        """
        baseDir = natqh.getUserDirectory()
        userDir = natqh.getUnimacroUserDirectory()
        originalUnimacroDir = natqh.getOriginalUnimacroDirectory()
        sampleBases = [baseDir.lower()]
        if originalUnimacroDir.lower() not in sampleBases:
            sampleBases.append(originalUnimacroDir.lower())
        if not os.path.isdir(commandDir):
            try:
                os.mkdir(commandDir)
            except OSError:
                print 'cannot make inifiles directory: %s'% commandDir
                return
        trySampleDirs = [os.path.join(sample, 'sample_ini', self.language +"_inifiles") for
                         sample in sampleBases]
        sampleDirs = [sample for sample in trySampleDirs if os.path.isdir(sample)]
        if self.language != 'enx':
            enxSampleDirs = [os.path.join(sample, 'sample_ini', "enx_inifiles") for
                 sample in sampleBases]
            sampleDirs.extend([sample for sample in enxSampleDirs if os.path.isdir(sample)])

        #print 'sampleDirs: %s'% sampleDirs
        inifile = os.path.join(commandDir, fileName)
        if sampleDirs:
            tryInifilesamples = [os.path.join(sample, fileName) for
                                    sample in sampleDirs]
            inifileSamples = [sample for sample in tryInifilesamples
                              if os.path.isfile(sample)]
        else:
            inifileSamples = None
        if not os.path.isfile(inifile):
            if inifileSamples:
                sample = inifileSamples[0]
                print 'take sample inifile: %s'% sample
                shutil.copyfile(sample, inifile)
            else:
                print 'could not find a valid sample inifile "%s" in directories: %s'%\
                        (fileName, sampleDirs)

    def TryToMakeDefaultInifile(self, commandDir, inifileName, language):
        """must be checked"""
    
        userDir = natqh.getUnimacroUserDirectory()
        modName = self.__module__

        if self.language != 'enx':
            enxFolder =  os.path.normpath(os.path.join(userDir, "enx_inifiles"))
            if os.path.isdir(enxFolder):
                enxVersion = os.path.join(enxFolder, modName + '.ini')
                if os.path.isfile(enxVersion):
                    self.makeDefaultInifile(enxVersion=enxVersion)
                else:
                    self.makeDefaultInifile()
        else:
            self.makeDefaultInifile()

     
     
                
    def checkInifile(self):
        """checking the inifile for changes

        Initialisation is supposed to have been done in the routine
        startInifile.
        """
        newDate = natqh.getFileDate(self.inifile)

        if newDate == 0:
            return # error, no inifile active

        if newDate > self.iniFileDate:
            oldName = self.ini.get('grammar name', 'name')
            #print 'newDate of %s'% self.inifile
            self.iniFileDate = newDate
            print '---(re)loading inifile: %s'% self.inifile ##, self.iniFileDate)
            try:
                self.ini = inivars.IniVars(self.inifile, repairErrors=1)
                self.fillInstanceVariables()
                self.switchOn()
                self.openedInifile = 0
                newName = self.ini.get('grammar name', 'name')
                #
                # NOTE: here we rely on the definition of the grammar specification as gramSpec class variable!
                #
                previousGramSpec = copy.copy(self.gramSpec)
                if isinstance(self.gramSpec, basestring):
                    previousGramSpec = self.gramSpec
                elif type(self.gramSpec) == types.ListType:
                    previousGramSpec = '\n'.join(self.gramSpec)
                else:
                    raise TypeError('At reload inifile of IniGrammar "%s", gramSpec should be a string or a list, not: %s'% (self.name, type(self.gramSpec)))
                translated = self.translateGrammar(self.originalGramSpec)  # pass originalGramSpec...
                if translated and translated != previousGramSpec:
                    print 'translated: %s'% translated
                    print 'previous: %s'% previousGramSpec
                    self.gramSpec = translated
                    fullPath = natlinkmain.loadedFiles[self.__module__]
                    natqh.setCheckForGrammarChanges(1)
                    print 'going to reload grammar %s (full path: %s)'% (self.name, fullPath)
                    os.utime(fullPath, None)
                    self.DisplayMessage('grammar %s will be reloaded at next utterance'% self.name)
                    self.iniFileDate = natqh.getFileDate(self.inifile)  # just in case it has been changed during translate
                #elif translated:
                #    print 'translation identical, no reload necessary for %s'% self.name

            except inivars.IniError:
                mess = str(sys.exc_info()[1])
                if not self.openedInifile:
                    self.openFileDefault(self.inifile)
                self.openedInifile = 1

                if notifyIniErrors:
                    self.message(mess)
                print '****error inifile: '+mess
                return
            return 1  # signalling things changed

    
    def openFolderDefault(self, foldername, mode=None, windowStyle=None):
        """open the folder in the default window"""
        mode = mode or 'open'
        if not os.path.isdir(foldername):
            self.DisplayMessage('the folder you want to open does not exist: %s'% foldername)
            return
        return UnimacroBringUp(app=None, filepath=foldername)
        #
        #
        #print 'open with AppBringup: %s'% foldername
        #natlink.execScript('AppBringup "%s"'% foldername)
        ##win32api.ShellExecute(0, mode, foldername, '', '', windowStyle or win32con.SW_SHOWNORMAL)
        ##int = ShellExecute(hwnd, op , file , params , dir , bShow )
        ##natqh.AppBringUp('folder', foldername, windowStyle=windowStyle)



    def openWebsiteDefault(self, website, openWith=None):
        """open the website in the default window or specific openWith program

        
        """
        print 'openWebsiteDefault: %s (openWith: %s)'% (website, openWith)
        appname = "website %s"% website
        if openWith:
            return UnimacroBringUp(openWith, website)
        else:
            return UnimacroBringUp(None, website)

    def openFileDefault(self, filename, mode=None, windowStyle=None, name=None, openWith=None):
        """open the file in the default window

        if the file is opened with AppBringUp ( not one of the exceptions), the optional
        name is given, which can be used with to switch to command!
        In the messages window the name of the AppBringUp command is given.
        
        """
        #print 'openFileDefault: %s (mode: %s, openWith: %s)'% (filename, mode, openWith)
        appname = "file %s"% filename
        mode = mode or 'open'
        
        if not os.path.isfile(str(filename)):
            self.DisplayMessage('the file you want to open does not exist: %s'% filename)
            return
        if openWith:
            return UnimacroBringUp(openWith, filename)
        elif mode:
            return UnimacroBringUp(mode, filename)
        else:
            return UnimacroBringUp(None, filename)

        #    
        #    #dummy, extension = os.path.splitext(filename)
        #    #openWith = getAppForEditExt(extension)
        #    #if openWith == 'edit' and extension in  ['.ini']:
        #    #    openWith = status.getUnimacroIniFilesEditor()
        #    #    if not openWith:
        #    #        openWith = "notepad"
        #        
        #if openWith:
        #    if openWith == 'notepad':
        #        Prog = os.path.join(natlinkcorefunctions.getExtendedEnv('SYSTEM'), 'notepad.exe')
        #    else:
        #        Prog = getAppPath(openWith) or getAppName(openWith) 
        #    natqh.AppBringUp(appname, Exec=Prog, Args=filename,
        #                     windowStyle=windowStyle)
        #        #win32api.ShellExecute(0, 'open', editProg, filename, "", 1)
        #    return
        #else:
        #    natqh.AppBringup(appname, Exec=None, Args=filename,
        #                     windowStyle=windowStyle)

    def getFromInifile(self, words, section, noWarning=0):
        """extract value from inifile, with words as possible keys


        If "words" is a string (single word from the list of words),
        only this word is looked up, returns '' if not found

        If "words" is a list, the first hit is returned.        
        """
        if self.ini == None:
            self.error('no valid inifile')
        if type(words) == types.StringType:
            v = self.ini.get(section, words)
            return v

        # list of words, normal case:
        for w in words:
            v = self.ini.get(section, w, None)
            if v != None:
                return v
        else:
            if not noWarning:
                print 'warning getFromInifile: no value found, ' \
                      'words: %s, section: %s, module: %s' % \
                              (words, section, self.__module__)
            return ''

    def getName(self):
        if "name" in dir(self):
            return self.name
        
        n = self.__module__
        if n[0] == "_":
            n = n[1:]
        if n.find("_") > 0:
            n = n.replace("_", " ")
        return n
    GetName = getName   # consistency with Bart Jan

    def GetDictionaries(self):
        inisections = self.ini.get()
        d = {}
        for listname in self.validLists:
            dictname = '%sDict'% listname
            if dictname in dir(self):
                D = getattr(self, dictname)
                if type(D) == types.DictType:
                    d[listname] = D.copy()
            elif listname in inisections:
                D = {}
                for key in self.ini.get(listname):
                    D[key] = self.ini.get(listname, key)
                d[listname] = D.copy()
        return d

    def makeDefaultInifile(self, inifile=None, enxVersion=None):
        """initialize as a starting example the inifile

        This default only sets in the section "general"  the key
        "initial on" to 0 or off (so grammar is initial off)         

        """       
        inifile = inifile or self.inifile
        print '----------making new default inifile %s'% inifile
        if enxVersion:
            t = open(enxVersion, 'r').read()
            open(inifile, 'w').write(t)
            
            self.ini = inivars.IniVars(inifile, repairErrors=1)
            self.ini.delete('grammar name')
            self.ini.delete('grammar words')
            self.ini.set('grammar name', 'name', '')
            self.ini.set('an instruction', 'intro', 'The english inifile is taken as example for this new language inifile for "%s"'% self.language)
            self.ini.set('an instruction', 'note 1', 'Translate [grammar name] and [grammar words] (right side of =)')
            self.ini.set('an instruction', 'note 2', 'Prevent quotes in names')
            self.ini.set('an instruction', 'note 3', 'Translate words in other sections as appropriate, see also Unimacro website...')
            self.ini.set('an instruction', 'note 4', 'If this is a new language for existing grammar files,')
            self.ini.set('an instruction', 'note 5', 'please contact the Unimacro developers to publish your translation in future versions!')
            self.ini.set('an instruction', 'note 6', 'This section can be deleted after reading')
        else:
            open(inifile, 'w').write('\n')
            self.ini = inivars.IniVars(inifile, repairErrors=1)
            self.ini.set('general', 'initial on', '1')
            self.ini.set('an instruction', 'intro', 'There was no sample/example inifile for this grammar.')
            self.ini.set('an instruction', 'note 1', 'Please fill in your own sections if relevant')
            if self.language == 'enx':
                self.ini.set('an instruction', 'note 2', 'If this is a grammar useful to others, please consider publishing it in the Unimacro area,')
                self.ini.set('an instruction', 'note 3', 'you can ontact the Unimacro developers through the Unimacro website for this.')
                self.ini.set('an instruction', 'note 4', 'This section can be deleted after reading')
            else:
                self.ini.set('an instruction', 'note 2', 'Please consider to write your grammar in English first and then make a translation to "%s"'% self.language)
                self.ini.set('an instruction', 'note 3', 'This section can be deleted after reading')
        
        self.fillDefaultInifile(self.ini)
        self.ini.close()

    def fillDefaultInifile(self, ini):
        """set initial settings for ini file, overload!

        """
        print '----------standard filling of default ini file for module: %s'% self.__module__
        ini.set('general', 'initial on', '1')
        name = self.__module__.strip('_')
        name = name.replace("_", " ")
        ini.set('grammar name', 'name', name)
        for l in self.validLists:
            if not('iniIgnoreGrammarLists' in dir(self) and l in self.iniIgnoreGrammarLists):
                ini.set(l)
                
    def error(self, message):
        """gives an error message, and leaves variable Error

        currently prints a message, and switches off the grammar
        """
        print '---- error in module %s: %s'% (self, message)
        raise UnimacroError('error %s in module %s: %s'% \
                            (sys.exc_info()[0], self.GetName(), message))
                            

        
    def removeFromList(self, L, toRemove):
        """removes in place items from list, calls error routine if missing things

        returns nothing!, list L list changed in place
             """
        if type(L) != types.ListType:
            self.error('not a list in "removeFromList": %s'% L)
            return

        if not L:
            if not toRemove:
                return
            else:
                self.error('removedFromList, list is empty, but toRemove is not empty: %s'% toRemove)
                return
        if not toRemove:
            return
        elif type(toRemove) == types.StringType:
            try:
                L.remove(toRemove)
            except:
                self.error('removeFromList, item to remove is not in list: %s'% toRemove)
                return
        elif type(toRemove) in (types.ListType, types.TupleType):
            for r in toRemove:
                try:
                    L.remove(r)
                except:
                    self.error('removeFromList, item to remove is not in list: %s'% r)
                    return
        else:
            self.error('removeFromList, invalid type for variable "toRemove": %s'% toRemove)
            return
        

    ## Here are the grammar rules, that interpret a given number:
    def gotResults___1to9(self,words,fullResults):
        numWords = self.getNumbersFromSpoken(words, asStr=1)
        self._sofar = self._sofar + ''.join(numWords)
    gotResults___0to9 = gotResults___1to9        

    def gotResults___1to99(self,words,fullResults):
        numWords = self.getNumbersFromSpoken(words, asStr=1)
        #print '1to99, words: %s, sofar: %s, add: %s'% (words, self._sofar, numWords)
        if self.language == 'nld':
            self._sofar = self._sofar + ''.join(numWords)
        else:
            # make '40 3' -->> '43' etc:
            L = []
            lenWords = len(numWords)
            for w in numWords:
                if len(w) == 1:
                    if L and int(L[-1]) and L[-1][-1] == '0':
                        print 'scratch away one 0, before: %s'% L
                        L[-1] = L[-1][:-1]  ## scratch away one 0, 40 3 must become 43, so list element '40' becomes '4'
                        print 'scratch away one 0, after: %s'% L
                L.append(w)
                            
            self._sofar = self._sofar + ''.join(L)
        #print 'sofar after rule: %s'% self._sofar

    gotResults___0to99 = gotResults___1to99
    gotResults___10to99 = gotResults___1to99


    def gotResults___hundreds(self,words,fullResults):
        assert not self._hundreds
        if self._sofar:
            self._hundreds = int(self._sofar)*100
        else:
            self._hundreds = 100
        self._sofar = ''

    def gotResults___thousands(self,words,fullResults):
        assert not self._thousands
        if self._sofar:
            sofar = int(self._sofar)
            if self._hundreds:
                sofar = self._hundreds + sofar
            
            self._thousands = sofar*1000
        elif self._hundreds:
            self._thousands = self._hundreds*1000
        else:
            self._thousands = 1000
        self._sofar = ''
        self._hundreds = 0
        

    def gotResults___millions(self,words,fullResults):
        assert not self._millions
        if self._sofar:
            sofar = int(self._sofar)
            if self._hundreds:
                sofar = self._hundreds + sofar
            elif self._thousands:
                sofar = self._thousands + sofar
                
            self._millions = sofar*1000000
        else:
            self._millions = 1000000
        self._sofar = ''
        self._hundreds = self._thousands = 0

    def gotResults_float(self,words,fullResults):
        for w in words:
            if self.hasCommon(words, ['comma', 'komma']):
                self._decimal = ','
            elif self.hasCommon(words, ['point', 'dot', 'punt']):
                self._decimal = '.'
            else:
                raise UnimacroError('no valid word for float separator found: %s'% words)
            self.collectNumber(part=1) # before the , or .


    def gotResults_integer(self,words,fullResults):
        pass  # only having subrules... Edit, number

    def waitForNumber(self, name):
        """set up variable name "name", to be filled with a number

        also initialise the necesarry instance variables
        """        
        self._waitForNum = name
        self._integerPart = self._millions = self._thousands = self._hundreds = 0
        self._sofar = ''
        self._decimal = ''
        # just in case it is called in the grammar logic:
        setattr(self, name, None)
        
        
    def collectNumber(self, part=0, asNumber=0):
        # if you're waiting for a number get it, and reset the parts
        if '_waitForNum' in dir(self) and self._waitForNum:
            N = self._millions
            N += self._thousands
            N += self._hundreds
            if N:
                if self._sofar:
                    N += int(self._sofar)
                # losing leading zeros
                Nstr = str(N)
            elif self._sofar:
                N = int(self._sofar)
                Nstr = self._sofar
            else:
                return # nothing to be done (yet)

            if part:
                # collect the first part of the number:
                self._integerPart = N
            else:
                if '_integerPart' in dir(self) and self._integerPart:
                    # floatingpoint numbers always return in string:
                    setattr(self, self._waitForNum, '%s%s%s'% (self._integerPart, self._decimal, Nstr))
                else:
                    # do the integer number:
                    if asNumber:
                        setattr(self, self._waitForNum, N)
                    else:
                        setattr(self, self._waitForNum, Nstr)
                        
                self._integerPart = 0
                self._waitForNum = ''
                
            # reset the whole bunch:
            self._millions = self._thousands = self._hundreds = 0
            self._sofar = ''
            
        #### end of number things        

    #### search things:==============================================

    def stopSearch(self, progInfo=None):
        """action after the search"""
        if not progInfo:
            progInfo = natqh.getProgInfo()
        if beforeOrAfter == 'before':
            if lastSearchDirection == 'up':
                s = '<<leftafterbacksearch %s>>'% len(lastSearchText)
            else:
                s = '<<leftafterforwardsearch %s>>'% len(lastSearchText)
        elif beforeOrAfter == 'after':
            if lastSearchDirection == 'up':
                s = '<<rightafterbacksearch %s>>'% len(lastSearchText)
            else:
                s = '<<rightafterforwardsearch %s>>'% len(lastSearchText)
        else:
            return
        print 'doing action in stop search'
        action(s, progInfo=progInfo)
            
    def getLastSearchDirection(self):
        """get global var lastSearchDirection"""
        return lastSearchDirection
    def setLastSearchDirection(self, direc):
        """set global var lastSearchDirection"""
        global lastSearchDirection
        if direc not in ['up', 'down']:
            raise UnimacroError('invalid search direction: "%s"'% direc)
        lastSearchDirection = direc
    

    def searchForText(self, direction, text=None, progInfo=None, modInfo=None,
                      beforeafter=None, insert=None, extend=None):
        """search up or down to text, invoke actions, for program differences

        if text == None, previous text is searched for
        beforeafter: None == ignore,  before1, after=2, leave after the selection.
        """
        global lastSearchText, lastSearchDirection, beforeOrAfter
        # pass progInfo to the actions, to keep them from changing inside the stuff:
        if progInfo == None:
            progInfo = natqh.getProgInfo(modInfo)
        prog, title, toporchild, windowHandle = progInfo
        if prog == 'excel':
            connectExcel(progInfo)
        elif prog == 'winword':
            connectWinword(progInfo)
        else:
            disconnectExcelOrWord()

        if text:
            if extend:
                lastSearchText = lastSearchText + text
            elif insert:
                lastSearchText = text + lastSearchText
            else:
                lastSearchText = text
                
        elif not lastSearchText:
            res = action("<<searchalwayscontinue>>", progInfo=progInfo)
            if not res:
                self.DisplayMessage('search command, but no search text yet')
                # if call came from generic_movement, in exclusive mode, cancel now:
                self.cancelMode()
                return
        if beforeafter:
            if beforeafter == 'for':
                beforeOrAfter = None
            elif beforeafter in ('before', 'after'):
                beforeOrAfter = beforeafter
            else:
                self.DisplayMessage('search command, beforeafter has invalid value: %s'% beforeafter)
                # if call came from generic_movement, in exclusive mode, cancel now:
                self.cancelMode()
                return
                

        # remember the last direction:    
        lastSearchDirection = direction
        if text:
            if appProgram == 'excel':
                if direction == 'up':
                    sdir = 2
                else:
                    sdir = 1
                ac = app.ActiveCell
                sheet.Cells.Find(What=lastSearchText,After=ac, SearchDirection=sdir).Activate()
               
            # search with new text:
            elif direction == 'up':
                lastSearchDirection = direction
                res = action("<<startbacksearch>>", progInfo=progInfo)
                if res:
                    # back search implemented in this program:
                    keystroke(text)
                    action("<<searchgo>>", progInfo=progInfo)
                else:
                    # 
                    self.DisplayMessage('search from top for "%s"'% text)
                    action("<<documenthome>>", progInfo=progInfo)
                    self.searchForText('down', text)
                    return -1  # changed direction
            else:
                action("<<startsearch>>", progInfo=progInfo)
                keystroke(lastSearchText, progInfo=progInfo)
                action("<<searchgo>>", progInfo=progInfo)
        else:
            # continue search with previous text:
            if appProgram == 'excel' and lastSearchText:
                if direction == 'up':
                    sdir = 2
                else:
                    sdir = 1
                ac = app.ActiveCell
                sheet.Cells.Find(What=lastSearchText,After=ac, SearchDirection=sdir).Activate()
            elif direction == 'up':
                if beforeOrAfter == 'after':
                    action("<<leftbeforebacksearch %s>>"% len(lastSearchText), progInfo=progInfo)
                res = action("<<searchback>>", progInfo=progInfo)
                if not res:
                    # back search failed, do forward search from top
                    self.DisplayMessage('search from top')
                    print 'start in top instead of back search'
                    action("<<documenthome>>", progInfo=progInfo)
                    self.searchForText('down', progInfo=progInfo)
                    return -1
            else:
                if beforeOrAfter == 'before':
                    action("<<rightbeforeforwardsearch %s>>"% len(lastSearchText), progInfo=progInfo)
                res = action("<<searchforward>>", progInfo=progInfo)

        # should return 0 if all was well, -2 if fail (-1 is for change direction)
        fail = self.searchFailed(progInfo=progInfo)
        if fail == -2:
            self.cancelMode()
            action("<<searchfailed>>", progInfo=progInfo)
        return fail            

    def searchFailed(self, progInfo=None):
        """signal if search was fail, invalid string or end of document"""
        if not progInfo:
            return # no information
        # old info:
        prog, title, toporchild, windowHandle = progInfo
        nprogInfo = natqh.getProgInfo() # for checking if window or title changed
        if (prog == 'natspeak' and title.find('dragonpad') >= 0) or \
           (prog == 'notepad' and title.find('notepad') >= 0) or \
           (prog == 'iexplore'):
            # hit if title changes, notepad 
            if progInfo != nprogInfo:
                print '%s: window changed, cancel search'% prog
                return -2 # cancelMode, because window title changed
        elif (prog == 'winword' and toporchild == 'top' and nprogInfo[2] == 'child'):
            print '%s: search failed, cancel search'% prog
            return -2 # cancelMode, because window title changed

    def searchMarkSpot(self, progInfo=None):
        """Mark the place where you can return with searchGoBack"""
        global comingFrom
        if progInfo == None:
            progInfo = natqh.getProgInfo()
        prog, title, toporchild, windowHandle = progInfo
        if prog == 'excel':
            connectExcel(progInfo)
            ac = app.ActiveCell
            comingFrom = ac.Address
        elif prog == 'winword':
            connectWinword(progInfo)
            if appProgram == 'winword':
                doc.Bookmarks.Add("unimacrosearch")
                print 'search grammar, word bookmark set'

    def searchGoBack(self, progInfo=None):
        """go back to previous place, excel or word"""
        if progInfo == None:
            progInfo = natqh.getProgInfo()
        prog, title, toporchild, windowHandle = progInfo
        if prog == 'excel':
            connectExcel(progInfo)
        elif prog == 'winword':
            connectWinword(progInfo)
        else:
            # eg ultraedit knows this:
            action("<<searchgoback>>")
        
        if appProgram == 'excel':
            print 'search go back to: "%s"'% comingFrom
            sheet.Range(comingFrom).Select()
        elif appProgram == 'winword':
            # What -1 = wdGotoBookmark:
##            print 'doc: %s'% doc
##            print 'doc.range: %s'% doc.Range
##            doc.Range().GoTo(-1,Name="unimacrosearch")
            print 'search go back does not work yet'

DocstringGrammarAncestor=IniGrammar    
class DocstringGrammar(DocstringGrammarAncestor):
    """grammar base which reads the rules from the docstrings of the
    rule functions
    """
    __inherited=DocstringGrammarAncestor
    nIndentSubrule = 4
    giveFullResults = 0
    def __init__(self):
        self.ruleFuncsDict = {} # records the functions corresponding to rules
        gramSpecFromDocstring = self.buildGramSpecFromDocstrings() or ""
        try:
            classGramSpec = self.__class__.gramSpec
            if not isinstance(classGramSpec, basestring):
                classGramSpec = '\n'.join(classGramSpec)
        except AttributeError:
            classGramSpec = ""

        totalgramSpec = gramSpecFromDocstring + '\n' + classGramSpec
        self.gramSpec = totalgramSpec.strip()
        self.originalGramSpec = self.gramSpec
        self.__inherited.__init__(self)

    def buildGramSpecFromDocstrings(self):
        L = []
        posfunc = []
        for item in dir(self):
            parts = item.split("_", 1)
            if len(parts) == 2 and parts[0] in ( 'rule', 'subrule', 'importedrule'):
                func = getattr(self, item)
                #print 'item: %s, type: %s'% (item, type(func))
                if type(func) == types.MethodType:
                    docstring = func.__doc__
                    if docstring is None or not docstring.strip():
                        if parts[0] in ('rule', 'subrule'):
                            raise ValueError('(sub)rule %s must have a docstring'% item)
                    lineno = func.func_code.co_firstlineno
                    posfunc.append( (lineno, item) )
        #print 'posfunc: %s'% posfunc
        if not posfunc:
            return ''
        posfunc.sort()
        #print 'posfunc: %s'% posfunc
        for dummy, item in posfunc:
            parts = item.split("_", 1)
            rulename = parts[1]
            func = getattr(self, item)
            docstring = func.__doc__ or ''
            ruledef = self.getRuleDefFromDocstring(docstring, rulename, parts[0])
            self.ruleFuncsDict[rulename] = func
            L.append(ruledef)
        #print 'ruleFuncsDict: %s'% self.ruleFuncsDict.keys()
        return '\n'.join(L)

    def getRuleDefFromDocstring(self, doc, rulename, typeOfRule):
        L = []
        
        exported = typeOfRule == 'rule'
        imported = typeOfRule == 'importedrule'
        doclines = doc.split('\n')
        doclines = self.adjustSpacingOfDocList(doclines)
        hadStart = 0
        lastLine = None
        rulePart = ''
        lenRulePart = 0
        for i, line in enumerate(doclines):
            spaces = ' '*(len(line) - len(line.lstrip()))
            line = line.strip()
            if line == '' or line.startswith('#'):
                L.append(line)
                continue
            if hadStart:
                L.append(' '*lenRulePart + spaces + line)
                lastLine = i
            else:
                if exported:
                    rulePart = '<%s> exported = '% rulename
                    lenRulePart = len(rulePart)
                    rule = spaces + rulePart + line
                else:
                    rulePart = '<%s> = '% rulename
                    lenRulePart = len(rulePart)
                    rule = spaces + rulePart + line
                hadStart = 1
                lastLine = i
                #print 'item:', rulename, ' rule: ', rule
                L.append(rule)
        if lastLine is None:
            if imported:
                rule = '<%s> imported;'% rulename
                L.append(rule)
            else:
                raise ValueError('no ruledefinition found for rule: %s'% rulename)
        else:
            if imported:
                raise ValueError('imported rule may have no ruledefinition: %s'% rulename)
            else:
                L[lastLine] = self.addSemicolonToRule(L[lastLine])
            
        if typeOfRule == 'subrule' and self.nIndentSubrule:
            spaces = ' '*self.nIndentSubrule
            L = [spaces+l for l in L]
        return '\n'.join(L)

    def addSemicolonToRule(self, line):
        """add a semicolon after the line, before comments
        """
        parts = line.split("#", 1)
        if len(parts) == 1:
            line = line.rstrip()
            if line.endswith(';'):
                return line
            else:
                return line + ';'
        else:
            rulePart = parts[0].rstrip()
            nSpaces = len(parts[0]) - len(rulePart)
            if rulePart.endswith(';'):
                return line
            else:
                spaces = ' '*(nSpaces)
                return rulePart + ";" + spaces + "#" + parts[1]
                
    def adjustSpacingOfDocList(self, docList):
        """remove consistent leading spaces from subsequent lines
        also remove empty first and/or empty last line
        """
        lenLeader = 11
        if not docList:
            return docList
        nSpaces = [len(d) - len(d.lstrip()) for d in docList]
        nSpacesLineOne = nSpaces[0]
        del nSpaces[0]
        if not nSpaces:
            return [docList[0].strip()]
        if len(nSpaces) > 0 and not docList[-1].strip():
            del nSpaces[-1]
            del docList[-1]
        if not nSpaces:
            return [docList[0].strip()]
           
        minSpaces = min(nSpaces)
        
        if minSpaces >= nSpacesLineOne + lenLeader:
            # only strip the leader part for nice presentation:
            D = []
            nToStrip = nSpacesLineOne + lenLeader
            for i, d in enumerate(docList):
                if i:
                    D.append(d[nToStrip:])
                else:
                    D.append(d)
        else:
            # strip according to minSpaces, strip also item 0
            D = []
            for i, d in enumerate(docList):
                if i:
                    d = d[minSpaces:]
                    D.append(d)
                else:
                    D.append(d.lstrip())
        return D
    
    def callRuleResultsFunctions(self, seqsAndRules, fullResults):
        """call the rule functions, can be overloaded (eg in DocstringGrammar)
        
        Also give self.nextRule (the name) self.nextWords, self.prevRule, self.prevWords
        so the result of the adjacent rules are known
        
        the new style functions prevail over the gotResults_ functions, ie
        the latter are only visited if the former fail.
        """
        ruleName, ruleWords = None, None
        for i, x in enumerate(seqsAndRules):
            if i == 0:
                ruleName, self.nextRule = None, x[1]
                ruleWords, self.nextWords = [], x[0]
            else:
                self.prevRule, ruleName, self.nextRule = ruleName, self.nextRule, x[1]
                self.prevWords, ruleWords, self.nextWords = ruleWords, self.nextWords, x[0]
                self.doRuleIfExists( ruleName, ruleWords, fullResults)

        self.prevRule, ruleName, self.nextRule = ruleName, self.nextRule, None
        self.prevWords, ruleWords, self.nextWords = ruleWords, self.nextWords, []
        self.doRuleIfExists(ruleName, ruleWords, fullResults)

    def doRuleIfExists(self, ruleName, ruleWords, fullResults):
        """overload from natlinkutils, calls for rule_ or gotResults_
        """
        if ruleName in self.ruleFuncsDict:
            # docstring style:
            func = self.ruleFuncsDict[ruleName]
            func(ruleWords)
        else:
            # classic (gotResults) style:                
            self.callIfExists( 'gotResults_'+ruleName, (ruleWords, fullResults) )
        
    def __str__(self):
        return '<docstringgr: %s>'% self.GetName()


    
class TestGrammarBase(IniGrammar):
    """Test grammar for displaying only the rules

    """
    def testReport(self, t):
        """print to DragonPad, or to Message window

        """        
        dp = natqh.matchModule('natspeak', 'DragonPad')
        if dp:
            natut.playString(t+'\n')
        else:
            print t       


    def resultsCallback(self, wordsAndNums, resObj):
        # TEST, of alleen uitvoer in DEBUG-venster, of in 
        # het NatSpeak-venster.
        if type(wordsAndNums) == type(''): 
            recogType = wordsAndNums
        else:
            recogType = 'self'

        # do nothing more if the recog results were not for this grammar
        if type(wordsAndNums) != type([]):
            return None

        # we firnst convert the passed array of word/ruleNumbers into an
        # array of word/ruleNames and an array of only words
        words = []
        fullResults = []
        for x in wordsAndNums:
            words.append( x[0] )
            fullResults.append( ( x[0], self.ruleMap[x[1]] ) )

        seqsAndRules = []
        for x in fullResults:
            if len(seqsAndRules) > 0 and seqsAndRules[-1:][0][1] == x[1]:
                # duplicate rule, append previous entry
                seqsAndRules[-1:][0][0].append(x[0])
            else:
                seqsAndRules.append( ([x[0]], x[1]) )
        
        # now we make the callbacks (in each case we only call the fucntion 
        # if it exists in the derived)
        # - we first call gotResultsInit
        # - then we make one callback for each different rule found as we
        #   sequentially scan the results (see seqsAndRules example)
        # - finally we call gotResults
        #self.callIfExists( 'gotResultsInit', (words, fullResults) )
        self.testReport('intercepted======='+`words`)
        for x in seqsAndRules:
            # i.p.v. call naar gotResults_... wordt hier een debugregel afgedrukt:
            self.testReport(`x[1]`+" :\t"+`x[0]`)
        #self.callIfExists( 'gotResults', (words, fullResults) )
        self.testReport('-'*20)
        
#number things, to be called from IniGrammar and from a user grammar, gives an integer:
numberGrammar = {}
numberGrammar['nld'] = """
<__1to99> = {n1-99};
<__0to99> = {n0-99};
<__0to9> = {n0-9};
<__10to99> = {n10-99};
<__hundreds> = honderd | <__1to99> honderd | honderd <__1to99> | <__1to99> honderd <__1to99> ;
<__thousands> = duizend | (<__1to99>|<__hundreds>) duizend | duizend (<__1to99>|<__hundreds>) |
                   (<__1to99>|<__hundreds>) duizend  (<__1to99>|<__hundreds>);
<__millions> = miljoen | miljoen (<__thousands>|<__hundreds>|<__1to99>) | (<__hundreds>|<__1to99>) miljoen |
                (<__hundreds>|<__1to99>) miljoen (<__thousands>|<__hundreds>|<__1to99>) ;
<integer> =  <__0to9>+ |  <__10to99>|
             <__0to9><__10to99>[<__0to99>+] |
             <__hundreds> | <__thousands> | <__millions>;
<float> = <integer> (punt|komma) (<__10to99>|<__0to9>+);
"""

numberGrammar['enx'] = """
<__0to99> = {n0-19} | {n20-90} | {n20-90}{n1-9};
<__1to99> = {n1-19} | {n20-90} | {n20-90}{n1-9};

<__hundreds> = hundred | <__1to99> hundred | hundred <__1to99> | <__1to99> hundred <__1to99> ;
<__thousands> = thousand | (<__1to99>|<__hundreds>) thousand | thousand (<__1to99>|<__hundreds>) |
                   (<__1to99>|<__hundreds>) thousand  (<__1to99>|<__hundreds>);
<__millions> = million | million (<__thousands>|<__hundreds>|<__1to99>) | (<__hundreds>|<__1to99>) million |
                (<__hundreds>|<__1to99>) million (<__thousands>|<__hundreds>|<__1to99>) ;
<integer> = <__0to99>+ | <__hundreds> | <__thousands> | <__millions>;
<float> = <integer> (point|dot) <__0to99>+;
"""
numberGrammar['deu'] = numberGrammar['nld']

numberGrammarTill999 = {}
numberGrammarTill999['nld'] = """
<__1to99> = {number1to99};
<__0to99> = {number0to99};
<__0to9> = {n0-9};
<__10to99> = {n10-99};
<__hundreds> = honderd | <__1to99> honderd | honderd <__1to99> | <__1to99> honderd <__1to99> ;
<integer> = <__0to9>+ |  <__10to99>|
             <__0to9><__10to99> | <__hundreds>;
<float> = <integer> (punt|komma|dot) (<__10to99>|<__0to9>+);
"""
numberGrammarTill999['enx'] = """
<__0to99> = {n0-19} | {n20-90} | {n20-90}{n1-9};
<__1to99> = {n1-19} | {n20-90} | {n20-90}{n1-9};
<__hundreds> = hundred | <__1to99> hundred | hundred <__1to99> | <__1to99> hundred <__1to99> ;
<integer> = <__0to99>+ | <__hundreds>;
<float> = <integer> (point|dot) <__0to99>+;

"""
numberGrammarTill999['deu'] = numberGrammarTill999['nld']

numberGrammarTill999999 = {}
numberGrammarTill999999['nld'] = """
<__1to99> = {n1-99};
<__0to99> = {n0-99};
<__0to9> = {n0-9};
<__10to99> = {n10-99};
<__hundreds> = honderd | <__1to99> honderd | honderd <__1to99> | <__1to99> honderd <__1to99> ;
<__thousands> = duizend | (<__1to99>|<__hundreds>) duizend | duizend (<__1to99>|<__hundreds>) |
                   (<__1to99>|<__hundreds>) duizend  (<__1to99>|<__hundreds>);
<integer> = <__0to9>+ |<__10to99>|
             <__0to9><__10to99>[<__0to99>]  | <__hundreds> | <__thousands>;
<float> = <integer> (punt|komma) (<__10to99>|<__0to9>+);
"""

numberGrammarTill999999['enx'] = """
<__0to99> = {n0-19} | {n20-90} | {n20-90}{n1-9};
<__1to99> = {n1-19} | {n20-90} | {n20-90}{n1-9};
<__hundreds> = hundred | <__1to99> hundred | hundred <__1to99> | <__1to99> hundred <__1to99> ;
<__thousands> = thousand | (<__1to99>|<__hundreds>) thousand | thousand (<__1to99>|<__hundreds>) |
                   (<__1to99>|<__hundreds>) thousand  (<__1to99>|<__hundreds>);
<integer> = <__0to99>+ | <__hundreds> | <__thousands>;
<float> = <integer> (point|dot)  <__0to99>+;
"""
numberGrammarTill999999['deu'] = numberGrammarTill999999['nld']

#connect to programs:
def connectExcel(progInfo):
    """connect to excel and leave parameters in global vars"""
    global app, appProgram, sheet
    prog, title, toporchild, windowHandle = progInfo
    
    if prog == 'excel':
        if appProgram != 'excel' or not app:
            app = win32com.client.Dispatch('Excel.Application')
        if app:
            sheet=app.ActiveSheet
            appProgram = 'excel'
            print 'excel application collected'
        else:
            print 'excel not connected'
            app = None
    else:
        app = None
        appProgram = ''
    return app

def connectWinword(progInfo):
    """connect to word and leave parameters in global vars"""
    global app, appProgram, doc   
    prog, title, toporchild, windowHandle = progInfo
    
    if prog == 'winword':
        if appProgram != 'winword' or not app:
            app = win32com.client.Dispatch('Word.Application')
        if app:
            doc = app.ActiveDocument
            appProgram = 'winword'
            print 'winword application collected'
        else:
            print 'winword not connected'
            app = None
            appProgram = ''
    else:
        app = None
        appProgram = ''
    return app
       
def disconnectExcelOrWord():
    """release connection with one of these"""
    global app, appProgram, doc, sheet
    if app:
        app = None
        appProgram = ''
        sheet = None
        doc = None
 
def formatListColumns(List, lineLen = 70, sort = 0):
    """formats a list in columns

    taken from utilsqh, 2011/1/4 QH

    Uses a generator function "splitList", that gives a sequence of
    sub lists of length n.

    The items are separated by at least two spaces, if the list
    can be placed on one line, the list is comma separated

    >>> formatListColumns([''])
    ''
    >>> formatListColumns(['a','b'])
    'a, b'
    >>> formatListColumns(['foo', 'bar', 'longer entry'], lineLen=5)
    'foo\\nbar\\nlonger entry'
    >>> formatListColumns(['foo', 'bar', 'longer entry'], lineLen=5, sort=1)
    'bar\\nfoo\\nlonger entry'
    >>> print formatListColumns(['afoo', 'bar', 'clonger', 'dmore', 'else', 'ftest'], lineLen=20, sort=1)
    afoo      dmore
    bar       else
    clonger   ftest
    >>> print formatListColumns(['foo', 'bar', 'longer entry'], lineLen=20)
    foo   longer entry
    bar

    """
    if sort:
        List.sort(caseIndependentSort)
    s = ', '.join(List)

    # short list, simply join with comma space:
    if len(s) <= lineLen:
        return s

    maxLen = max(map(len, List))

    # too long elements in list, return "\n" separated string:
    if maxLen > lineLen:
        return '\n'.join(List)


    nRow = len(s)/lineLen + 1
    lenList = len(List)

    # try for successive number of rows:
    while nRow < lenList/2 + 2:
        lines = []
        for i in range(nRow):
            lines.append([])
        maxLenTotal = 0
        for parts in splitList(List, nRow):
            maxLenParts = max(map(len, parts)) + 2
            maxLenTotal += maxLenParts
            for i in range(len(parts)):
                lines[i].append(parts[i].ljust(maxLenParts))
        if maxLenTotal > lineLen:
            nRow += 1
        else:
            return '\n'.join(map(string.strip, map(string.join, lines)))
    else:
        # unexpected long list:
        return '\n'.join(List)


def splitList(L, n):
    """generator function that splits a list in sublists of length n

    taken from utilsqh, 2011/1/4 QH

    """
    O = []
    for l in L:
        O.append(l)
        if len(O) == n:
            yield O
            O = []
    if O:
        yield O

