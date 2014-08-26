
import re
import sys, copy
reNote = re.compile(r"""([a-g](?:as|es|is|s)?)  # the notename
                        ([,']*)                 # elevation (octave up/down                        
                        ([0-9]*[.]*)            # duration, including augmentation dots
                        (.*)$                   # rest (assume no spaces)
                    """, re.VERBOSE+re.UNICODE)

reRest = re.compile(r"""([rRsS])  # the name of the rest
                        ##([,']*)                 # elevation (octave up/down                        
                        ([0-9]*[.]*)            # duration, including augmentation dots
                       ### (.*)$                   # rest (assume no spaces)
                    """, re.VERBOSE+re.UNICODE)
reBackslashedWord = re.compile(r"""([\\]\w+)""")


class LyNote(object):
    def __init__(self, s):
        self.originalInput = s
        self.setVariables(s)
        
    def setVariables(self, s):
        """parse s and set self.note etc."""
        
        m = reNote.match(s)
        n = reRest.match(s)
        self.note = self.elevation = self.duration = ""
        rest = s
        self.additions = ""
        self.backslashed = None
        if m:
            self.note = m.group(1)
            self.elevation = m.group(2)
            self.duration = m.group(3)
            rest = m.group(4)
            if rest:
                self.additions, self.backslashed = self.orderRest(rest)
                # backslashed is a list of \melisma etc.
        elif n:
            self.note = n.group(1)
            self.duration = n.group(2)
            self.elevation = self.rest = ""
            
        else:
            # can be incomplete, add fake note in front and do again
            sFake = 'c' + s
            self.setVariables(sFake)
            self.note = ""


    def isNote(self):
        return ( self.note != "" )
    
    def __str__(self):
        if self.backslashed:
            backslashed = " " + ' '.join(self.backslashed)
        else:
            backslashed = ""
        return "%s%s%s%s%s"% (self.note, self.elevation, self.duration, self.additions, backslashed)
    
    def __repr__(self):
        if self.backslashed:
            backslashed = " " + ' '.join(self.backslashed)
        else:
            backslashed = ""
        return "lynote: %s%s%s%s%s"% (self.note, self.elevation, self.duration, self.additions, backslashed)
    
    def getElevation(self):
        """return as int the elevation
        , = -1, ,, = -2, ' = 1, '' = 2 and "" = 0
        """
        if self.elevation == "":
            return 0
        if self.elevation.find(",") >= 0:
            return -len(self.elevation)
        if self.elevation.find("'") >= 0:
            return len(self.elevation)
        raise ValueError('getElevation: no valid elevation: "%s" (note: "%s"'% (self.elevation, self))
    
    def setElevation(self, elevation):
        """set the elevation string from the numeric elevation value
        0 = "", -1 = ",", 2 = "''" etc
        """
        if elevation == 0:
            return ""
        if elevation > 0:
            return "'"*elevation
        if elevation < 0:
            return ","*-elevation
        
        
    
    def orderRest(self, rest):
        if rest is None:
            return "", None
        
        if rest.find('\\') >= 0:
            additions = ""
            backslashed = []
            restList = reBackslashedWord.split(rest)
            for item in restList:
                if not item:
                    continue
                if reBackslashedWord.match(item):
                    backslashed.append(item)
                else:
                    additions += item

            return additions, backslashed
        return rest, None
        
    
    
    def updateNote(self, note):
        if isinstance(note, basestring):
            note = LyNote(note)
        
        for attr in ['note', 'duration']:
            value = getattr(note, attr)
            if value:
                setattr(self, attr, value)
        if note.elevation:
            nElv = note.getElevation()
            orgElv = self.getElevation()
            newElv = orgElv + nElv
            self.elevation = self.setElevation(newElv)
            
        if note.additions:
            if self.additions:
                self.additions += note.additions
            else:
                self.additions = note.additions
        if note.backslashed:
            if self.backslashed:
                self.backslashed.extend(note.backslashed)
            else:
                self.backslashed = copy.copy(note.backslashed)
                    
                
    
if __name__ == '__main__':
    
    #for s in ["", "(", r"(\break\melisma", r"\break\(-^2"]:
    #    mrest = reBackslashedWord.split(s)
    #    print 's: %s, mrest: %s'% (s, mrest)
    
    #for s in ["g,8.", "a", "cis'("]:
    #    lyn = LyNote(s)   
    #    print 'note: "%s", elevation: "%s", duration: "%s", additions: "%s"'% (lyn.note, lyn.elevation, lyn.duration, lyn.additions)
    #    print 'input: "%s", str: "%s", repr: "%s"'% (s, lyn, repr(lyn))
    #    lyn.updateNote("a")
    #    print 'note updated to a: %s'% lyn
    #    lyn.updateNote(r"c8.\(")
    #    print 'note updated to c 8. and \\(: %s'% lyn
    print 'melisma: ============================================='
    for s in ["r2", r"g,8.\melisma", r"a\melisma"]:
        lyn = LyNote(s)
        print 'note: "%s", elevation: "%s", duration: "%s", additions: "%s"'% (lyn.note, lyn.elevation, lyn.duration, lyn.additions)
        print 'input: "%s", str: "%s", repr: "%s"'% (s, lyn, repr(lyn))
        lyn.updateNote("a")
        print 'note updated to a: %s'% lyn
        lyn.updateNote(r"c8.\(")
        print 'note updated to c 8. and \\(: %s'% lyn
        lyn = LyNote(s)   
        