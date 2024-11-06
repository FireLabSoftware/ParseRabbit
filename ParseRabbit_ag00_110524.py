#!/usr/bin/env python -i
########################
## ParseRabbit version ag00 11_05_24: Flexible code for Excel Generation from high throughput sequence data
##  General use
##   pypy ParseRabbit##  input=<Sequence_read.fastq.gz>  template=<input_oligonucleotide> 
## ParseRabbit is a tool for building Excel workbooks that list and provisionally parse reads from a sequencing run
## Each output worksheet consists of
##   a. an introductory header including specifics of the data and code used to generate the worksheet
##   b. A line-by-line listing of sequece species, sorted by incidence- each with a provisional parsing of the read
## Inputs are high throughput sequencing datasets (generally fastq.gz files).
## Output spreadsheets are .xlsx files using the very wonderful xlsxwrite Python unit
##    (Source and citation: John McNamara; https://pypi.org/project/XlsxWriter/)
## Sequences from a fastq.gz file are read a specific filter (set in the code) is on one line of the dataset
## The code and specific command line options are also provided on the introductory page in
## single informational cells.
## ParseRabbit uses an ancillary program (VSG_Module) and the current used code is provided in a cell.
## Finally, run details (Python version and time) are provided
## Thus all code and parameters are provided as part of the final Excel Workssheet product
## While the major output of ParseRabbit is the simple listing of species, a provisional parsing
## of the individual sequences is provided as a preliminary visualization. Importantly, this simple
## graphic is not the result of any experimental observation or intelligent prediction, but simply a
## display of one possible configuration.  The displayed configurations show the highest scoring duplex that
## can be formed between an input oligonucleotide and the observed extendion.  Scoring for this is
## very simple: a single point is awarded for each matched base pair, with the last base of the original
## input material ('MyTemplate') allowed to contribute to the count of matched bases.  The proceeding (-2)
## nucleotide and any nonspecified (N) bases in the input are given an arbitrary low value of 0.1 to allow
## eventual prioritization. Untemplated bases at the beginning of the extension are given a penalty of 2.0.
## The display is for provisional visualization only-- many complex products could be parsed
## in different ways, with no a priori indication of which is an accurate reflection of history.
## Beyond the simple graphic, the individual parsed segments are noted, and these are likewise simple
## results of running the attached code, with no expectation of accurate representation of any
## sequence.  Many sequence products (particularly those which result from multiple or complex
## priming events but also a subset of simple priming events) could plausibly be reflected by several
## different structures, and such structures may indeed be all accurate, all inaccurate, or somwhere in
## between in their ability to describe events leading to the observed product.
## As a specific caution, because the terminal bases in the template and randomized based in template
## pools are used to prioritize the possible parsing options, the parses should not be used to evaluate
## the ensemble frequency of alternative events that involve these bases (e.g. cis- versus trans- priming
## experiments). Thus such use has not been applied for these tables in our studies.
## 
## Note that ParseRabbit is active code in the sense that adjustments to the underlying filters and parsing
## rules generally involve direct modifications to the code (not just command line adjustments).
## Thus a key element for any reproduction or parallel application of code should use the code present in
## this spreadsheet and not the code from our GitHub or other repository
## 
## Addresses of archival code and run information are as follows on this worksheet
## This Introductory Description: Cell 1,3
## A worksheet-specific narrative: Cell 2,3
## ParseRabbit Code used for the specific worksheet: Cell 2,3
## VSG_Module Code: Cell 3,3
## Command Line Parameters (including file names and dates): Cell 4,3
## Run History: Cell 5,3

myLinker1 = 'TGGAATTCTC'
InputFiles1 = '*.fastq.gz' # (input=)
Template1 = 'default'
ExcelFileBase1 = 'default'

from VSG_ModuleFP import *
vCommand()

import gzip
from collections import Counter
import xlsxwriter
import glob
import re


def FileInfo1(FileID):
    if type(FileID)==str:
        Fn1=FileID
    else:
        Fn1=FileID.name
    s11=os.stat(Fn1)
    return ','.join([Fn1,
                     '#',
                      'Path='+os.path.abspath(Fn1),
                        'Size='+str(s11[6]),
                        'Accessed='+strftime("%m_%d_%y_at_%H_%M_%S",localtime(s11[7])),
                        'Modified='+strftime("%m_%d_%y_at_%H_%M_%S",localtime(s11[8])),
                        'Created='+strftime("%m_%d_%y_at_%H_%M_%S",localtime(s11[9])),
                        'FullInfo='+str(s11)])


def addRichText1(C,mR,mC,ws):
    if type(C)==str or type(C)==int:
        ws.write_rich_string(mR, mC, aa,str(C),aa,'',aa)
        return 1
    nbsp = chr(160)
    maxC = Counter()
    minC = Counter()
    myL = []
    for (c,r) in C:
        maxC[r] = max(maxC[r] , c)
        minC[r] = min(minC[r] , c)
    maxC0 = max(maxC.values())
    minC0 = min(minC.values())
    numlines = max(maxC.keys())+1
    myWidth = maxC0-minC0
    for r in range(numlines):
        mylinelen = 0 
        curF = oo
        nub = ''
        for c in range(minC0,maxC0+1):
            if (c,r) in C:
                if C[(c,r)][1]!=curF:
                    if nub:
                        myL.extend([curF,nub])
                    curF = C[(c,r)][1]
                    nub = ''                    
                nub += C[(c,r)][0]                    
            else:
                nub+=nbsp
        if r==numlines-1:
            myL.extend([curF,nub])
        else:
            myL.extend([curF,nub+'\n'])
    ws.write_rich_string(mR, mC, *myL,aa)
class segment():
    def __init__(self,b,e,s):
        self.b = b
        self.e = e
        self.s = s[b:e]
        if not(self.s):
            self.s=''
import xlsxwriter ## Install with Pip through PyPi, or https://github.com/jmcnamara. Thank you John Mcnamara!
    
def HeaderTranspose1(hT2):
    hT0 =   [['<!--','Output_Key','','-->']]
    hT0 +=    [['<!--','ColumnNumber','ColumnHeader','-->']]
    for iT2,nT2 in enumerate(hT2):
        nT2 = nT2.strip()
        hT0 += [['<!--',iT2,nT2,'-->']]
    return hT0
def placexlsx1(listoflist,firstrownum,worksheet,myFormat):
    '''place lines onto the worksheet based on the listoflist input, return the current line'''
    currownum = firstrownum
    for currow in listoflist:
        for col,va in enumerate(currow):
            worksheet.write(currownum,col,va,myFormat)
        currownum+=1
    return currownum
class dRNA():
    def __init__(self,s):
        ## t is the input RNA ('template'), s is the sequence of the recovered RNA
        self.seq = s    ## trimmed sequence
        self.cor = s[:myTLen1] ## core sequence (initial template-derived sequence)
        self.ext = s[myTLen1:] ## extension sequence beyond core
        self.count = 1 ## number of instances of this sequence
        self.parsed = False
        l1 = len(self.ext)
        l2 = len(self.cor)
        p2p = Counter()
        o2o = Counter()
        anticor = vantisense(self.cor)
        for i1,c1 in enumerate(self.ext):   ## i1, c1 are position and base in extension
            for i2,(c2,c3) in enumerate(zip(antitem1,anticor)): ## i2 is position in anti-template, c2 is base in antiinput, c3 is actual base in anti-read
                if c1==c3 or c2=='N':
                    if c1==c3:
                        reward = 1.0
                    else:
                        reward = 0.1
                    p2p[(i1,i2)] = p2p[(i1-1,i2-1)]+reward
                    if (i1-1,i2-1) in o2o:
                        o2o[(i1,i2)] = o2o[(i1-1,i2-1)]
                    else:
                        o2o[(i1,i2)] = i1,i2
                else:
                    p2p[(i1,i2)] = -2*i1
                if i1==0 and i2>0 and self.cor[-1]==anticor[i2-1]:
                    p2p[(i1,i2)] += 1
                    if i1==0 and i2>1 and self.cor[-2]==anticor[i2-2]:
                        p2p[(i1,i2)] += 0.1
        if len(p2p)>=2 and p2p.most_common()[0][1]>=3:
            ((i11,i12),n1),((i21,i22),n2) = p2p.most_common(2)
            if n1>=2: # and n1>n2:
                self.parsed = True
                ## n1 is length of duplex minus first base in s1 that is in hybrid
                ## i11 is last base in s1 that is in hybrid
                ## nb will be the number of bases in the duplex
                ## so i11-nb = start position in probe, n1 = nb-(i11-nb)
                ## hence nb1 = (n1+i11)//2
                nb = i11-o2o[(i11,i12)][0]+1 ## length of simply templated extension
                ps1 = 1+i11-nb
                pe1 = 1+i11
                ps2 = l2-i12-1
                pe2 = l2-i12-1+nb
                ns = 0 ## length of priming stem
                while (l2-ns-1>pe2+ns) and ((s[l2-ns-1],s[pe2+ns]) in cmp1):
                    ns += 1
                self.f = segment(0,ps2,s)
                self.t = segment(ps2,pe2,s)
                self.s = segment(pe2,pe2+ns,s)
                self.l = segment(pe2+ns,l2-ns,s)
                self.p = segment(l2-ns,l2,s)
                self.B = segment(l2,l2+ps1,s)
                self.C = segment(l2+ps1,l2+pe1,s)
                self.E = segment(l2+pe1,l2+l1,s)
        if not(self.parsed):
            self.f = segment(0,l2,s)
            self.t = segment(l2,l2,s)
            self.s = segment(l2,l2,s)
            self.l = segment(l2,l2,s)
            self.p = segment(l2,l2,s)
            self.B = segment(l1+l2,l1+l2,s)
            self.C = segment(l1+l2,l1+l2,s)
            self.E = segment(l1+l2,l1+l2,s)
    def diagramMeExcel(self):
        charC = Counter()
        if not(self.parsed):
            p = 0
            for i in range(len(self.cor)-1):
                if myTemplate1[i]=='N':
                    charC[(p,0)] = self.seq[i].lower(),nn
                else:
                    charC[(p,0)] = self.seq[i].lower(),oo
                p += 1
            i += 1
            charC[(p,1)] = self.seq[i].lower(),oo
            p += 0
            i += 1
            if i<len(self.seq):
                charC[(p,2)] = self.seq[i],oo
                if i<len(self.seq):
                    p -= 1
                    i += 1
                    while i<len(self.seq):
                        charC[(p,3)] = self.seq[i],oo
                        p -= 1
                        i += 1
            return charC            
        p = 0
        for i in range(self.f.b,self.f.e):
            if myTemplate1[i]=='N':
                charC[(p,0)] = self.seq[i].lower(),nn
            else:
                charC[(p,0)] = self.seq[i].lower(),ff
            p += 1
        for i in range(self.t.b,self.t.e):
            if myTemplate1[i]=='N':
                if (self.seq[i],self.seq[self.C.e-(i-self.t.b)-1]) in cmp1:
                    charC[(p,1)] = self.seq[i].lower(),nn
                else:
                    charC[(p,0)] = self.seq[i].lower(),nn
            else:
                charC[(p,1)] = self.seq[i].lower(),tt
            p += 1
        if self.B.s and not(self.s.s):
            displayLoop = self.l.s.lower()+self.B.s
            loopStart = i
            lenL = len(displayLoop)
            if lenL == 0:
                p -= 1
            elif lenL == 1:
                charC[(p,1)] = displayLoop[i-loopStart],ll
                if (p-1,1) in charC:
                    charC[(p-1,0)] = charC[(p-1,1)]
                    del(charC[(p-1,1)])
                i+=1
                p-=1
            elif lenL == 2:
                charC[(p,1)] = displayLoop[i-loopStart],ll
                i+=1
                p+=0
                charC[(p,2)] = displayLoop[i-loopStart],ll
                if (p-1,1) in charC:
                    charC[(p-1,0)] = charC[(p-1,1)]
                    del(charC[(p-1,1)])
                i+=1
                p-=1
            elif lenL%2==0:
                for k in range(lenL//2-1):
                    charC[(p,0)] =  displayLoop[i-loopStart],ll
                    i+=1
                    p+=1
                charC[(p,1)] =  displayLoop[i-loopStart],ll
                i+=1
                p+=0
                charC[(p,2)] =  displayLoop[i-loopStart],ll
                i+=1
                p-=1
                for k in range(lenL//2-1):
                    charC[(p,3)] =  displayLoop[i-loopStart],ll
                    i+=1
                    p-=1
            else:
                for k in range(lenL//2):
                    charC[(p,0)] =  displayLoop[i-loopStart],ll
                    i+=1
                    p+=1
                charC[(p,1)] = displayLoop[i-loopStart],ll
                i+=1
                p-=1
                for k in range(lenL//2):
                    charC[(p,2)] =  displayLoop[i-loopStart],ll
                    i+=1
                    p-=1
        else:
            for j in range(self.B.b,self.B.e):
                charC[(p,1)] = '-',oo
                p += 1
            for i in range(self.s.b,self.s.e):
                charC[(p,1)] = self.seq[i].lower(),ss
                p += 1
            i += 1
            lenL = len(self.l.s)
            if lenL == 0:
                p -= 1
            elif lenL == 1:
                charC[(p,1)] = self.seq[i].lower(),ll
                if (p-1,1) in charC:
                    charC[(p-1,0)] = charC[(p-1,1)]
                    del(charC[(p-1,1)])
                i+=1
                p-=1
            elif lenL == 2:
                charC[(p,1)] = self.seq[i].lower(),ll
                i+=1
                p+=0
                charC[(p,1)] = self.seq[i].lower(),ll
                if (p-1,1) in charC:
                    charC[(p-1,0)] = charC[(p-1,1)]
                    del(charC[(p-1,1)])
                i+=1
                p-=1
            elif lenL%2==0:
                for k in range(lenL//2-1):
                    charC[(p,0)] = self.seq[i].lower(),ll
                    i+=1
                    p+=1
                charC[(p,1)] = self.seq[i].lower(),ll
                i+=1
                p+=0
                charC[(p,2)] = self.seq[i].lower(),ll
                i+=1
                p-=1
                for k in range(lenL//2-1):
                    charC[(p,3)] = self.seq[i].lower(),ll
                    i+=1
                    p-=1
            else:
                for k in range(lenL//2):
                    charC[(p,0)] = self.seq[i].lower(),ll
                    i+=1
                    p+=1
                charC[(p,1)] = self.seq[i].lower(),ll
                i+=1
                p-=1
                for k in range(lenL//2):
                    charC[(p,2)] = self.seq[i].lower(),ll
                    i+=1
                    p-=1
            for i in range(self.p.b,self.p.e):
                charC[(p,2)] = self.seq[i].lower(),pp
                p -= 1
            for i in range(self.B.b,self.B.e):
                charC[(p,2)] = self.seq[i],BB
                p -= 1
        for i in range(self.C.b,self.C.e):
            charC[(p,2)] = self.seq[i],CC
            p -= 1
        for i in range(self.E.b,self.E.e):
            charC[(p,3)] = self.seq[i],EE
            p -= 1
        return charC
    

for myInputFileName1 in glob.glob(InputFiles1):
    if Template1 == 'default':
        EG73 = 'AGANNATTATTACGTGCTTTTGTTCAA'
        EG74 = 'TTNNACGTCAACGATATAAGTTTTGAC'
        ## Automated handling of the two oligos used for program development-- can be commented out for a more generic version
        if 'EG-73' in os.path.basename(myInputFileName1):
            myTemplate1 = EG73
        elif 'EG-74' in os.path.basename(myInputFileName1):
            myTemplate1 = EG74
        else:
            continue
    else:
        myTemplate1 = Template1
    antitem1 = vantisense(myTemplate1)
    myRE1 = re.compile('^'+myTemplate1.replace('N','.'))


    if ExcelFileBase1=='default':
        ExcelFileOutput1 = myInputFileName1.split('.')[0]+vnow+'_Parsed.xlsx'
    else:
        ExcelFileOutput1 = ExcelFileBase1+myInputFileName1.split('.')[0]+vnow+'_Parsed.xlsx'
    if not(ExcelFileOutput1.endswith('.xlsx')):
        ExcelFileOutput1 += '.xlsx'
    workBook1 = xlsxwriter.Workbook(ExcelFileOutput1)
    ##  a bunch of formats, not all used, for xcel cells
    ff = workBook1.add_format({'text_wrap': True, 'align':'right', 'font_name': 'Courier', 'font_size': 12, 'font_color':'black'})
    nn = workBook1.add_format({'text_wrap': True, 'align':'right', 'font_name': 'Courier', 'font_size': 12, 'font_color':'red'})
    tt = workBook1.add_format({'text_wrap': True, 'align':'right', 'font_name': 'Courier', 'font_size': 12, 'font_color':'black'})
    ss = workBook1.add_format({'text_wrap': True, 'align':'right', 'font_name': 'Courier', 'font_size': 12, 'font_color':'black'})
    ll = workBook1.add_format({'text_wrap': True, 'align':'right', 'font_name': 'Courier', 'font_size': 12, 'font_color':'black'})
    oo = workBook1.add_format({'text_wrap': True, 'align':'right', 'font_name': 'Courier', 'font_size': 12, 'font_color':'black'})
    pp = workBook1.add_format({'text_wrap': True, 'align':'right', 'bold':True, 'font_name': 'Courier', 'font_size': 12, 'font_color':'black'})
    BB = workBook1.add_format({'text_wrap': True, 'align':'right', 'font_name': 'Courier', 'font_size': 12, 'font_color':'black'})
    CC = workBook1.add_format({'text_wrap': True, 'align':'right', 'font_name': 'Courier', 'font_size': 12, 'font_color':'black'})
    EE = workBook1.add_format({'text_wrap': True, 'align':'right', 'font_name': 'Courier', 'font_size': 12, 'font_color':'black'})
    aa = workBook1.add_format({'text_wrap': True, 'align':'right'})
    # Formats for routine "housekeeping" in ParseRabbit Excel Documents
    DD = workBook1.add_format({'text_wrap': False, 'bold':True ,'align':'left','font_size': 14})  # Cell Content Identifier
    dd = workBook1.add_format({'text_wrap': True, 'bold':False ,'align':'left','font_size': 11})         
    hh = workBook1.add_format({'text_wrap': False, 'bold':False ,'align':'left','font_size': 12})        
    HH = workBook1.add_format({'text_wrap': False, 'bold':True ,'align':'left','font_size': 12})   
    zz = workBook1.add_format({'text_wrap': True, 'align':'left', 'font_name': 'Courier', 'font_size': 12, 'font_color':'black'})
    ZZ = workBook1.add_format({'text_wrap': True, 'align':'left', 'font_name': 'Courier', 'font_size': 12, 'font_color':'red'})

    workSheet1 = workBook1.add_worksheet()
    workSheet1.set_column(3,1,27)
    workSheet1.set_column(3,2,27)
    workSheet1.set_column(3,3,64)
    NarrativeDescription1 = ''
    for  L1 in open(sys.argv[0], mode='rt'):
        if L1.startswith('##') and L1[2]!='#':
            NarrativeDescription1 += L1.strip()[2:]+'||'        
        if L1[0]!='#': break


    Headers1 = ['Rank_'+myInputFileName1,]
    Headers1 += ['Sequence_',]
    Headers1 += ['Count_'+myInputFileName1,]
    Headers1 += ['PreliminarySimpleParse_'+sys.argv[0][:-3],]
    ftslpBCE1 = 'fray_seg template_seg stem_seg loop_seg primer_seg Bulge_seg Complement_seg Extra_seg'.split()
    for part1 in ftslpBCE1:
        Headers1 += [part1+'_seq',part1+'_begin',part1+'_end',part1+'_length']
            ##   f = fray (unpaired sequence at beginning of input oligo)  : black
            ##   t = template (sequence from input oligo that is copied)  : black
            ##   s = stem (landing pad in input oligo for primer end of sequence):  black
            ##   l = loop (unpaired region in input oligo):  black
            ##   p = primer (paired region at end of input oligo) : black
            ##   B = Bulge (unpaired region at beginning of extension) : gray 
            ##   C = Complement (first accepted paired region at beginning of extension) : red
            ##   E = Extra (any sequenmce after complementary stretch, may include matching sequences after a mismatch) - unmatched (gray) matched (red anti, blue sense)



    TaskHeader1 =  [['<!--ParseRabbit_Task_Header: '+ExcelFileOutput1+'-->',]]
    TaskHeader1 += [['<!--Command_Line: '+' '.join(sys.argv)+'-->',]]
    TaskHeader1 += [['<!--PythonVersion: '+','.join(sys.version.splitlines())+'-->',]]
    TaskHeader1 += [['<!--ParseRabbit_Version: '+FileInfo1(sys.argv[0])+'-->',]]
    TaskHeader1 += [['<!--RunTime: '+vnow+'-->',]]
    TaskHeader1 += [['<!--RunDirectory: '+os.getcwd()+'-->',]]
    TaskHeader1 += [['<!---->',]]
    AbbrevHeader1 = ''.join([x[0] for x in TaskHeader1])+'<!--ParseRabbitTableHeader-->'  ##ending with ':RabbitTableHeader' identifies a line as a row of table headers

    myRow1 = 0
    myRow1 = placexlsx1(TaskHeader1,myRow1,workSheet1,hh)
    myRow1 = placexlsx1(HeaderTranspose1(Headers1),myRow1,workSheet1,hh)
    myRow1 += 1
    myRow1 = placexlsx1([Headers1+[AbbrevHeader1,' ']],myRow1,workSheet1,HH)

        
        
    myTLen1 = len(myTemplate1)
    cmp1 = set([('G','C'),('C','G'),('A','T'),('T','A'),
                    ('G','N'),('C','N'),('A','N'),('T','N'),
                    ('N','C'),('N','G'),('N','T'),('N','A'),('N','N')])
    dC1 = Counter()  ## Keys are trimmed read sequence, values are parsed RNA or DNA objects (dRNA)

    ## 
    ##                          l l l
    ##  f f f t t t t t t s s s       l
    ##        C C C C C C p p p l l l
    ##    E E            B

    ##   f = fray (unpaired sequence at beginning of input oligo)
    ##   t = template (sequence from input oligo that is copied)
    ##   s = stem (landing pad in input oligo for primer end of sequence)
    ##   l = loop (unpaired region in input oligo)
    ##   p = primer (paired region at end of input oligo)
    ##   B = Bulge (unpaired region at beginning of extension)
    ##   C = Complement (first accepted paired region at beginning of extension)
    ##   E = Extra (any sequenmce after complementary stretch, may include matching sequences after a mismatch)
    ## for each of these components, we have a segment with three parameters
    ##  begin (.b), end (.e), and seq (.s).
    ##  .b is a zero-based position indicator for the start of the feature in the eventual read
    ##  .e is a zero-based position indicator for the first position in the read not covered by the feature
    ##  .s is the sequence of the feature

    for i1,L1 in enumerate(vOpen(myInputFileName1, mode='rt')):
        if i1%4!=1: continue
        if myRE1.search(L1) and (myLinker1 in L1):
            p1 = L1.rfind(myLinker1)
            s1 = L1[:p1]  ## captured insert
            if s1 in dC1:
                dC1[s1].count += 1
            else:
                dC1[s1] = dRNA(s1)
    for i1,s1 in enumerate(sorted(dC1.keys(), key= lambda x:-dC1[x].count)): ##lambda x:(dC1[x].seq[3:5],-dC1[x].count))):
        j1 = i1+myRow1
        d1 = dC1[s1]
        workSheet1.write(j1,0,i1+1)
        workSheet1.write(j1,2,d1.count)
        sL1 = []
        for ii1,c1 in enumerate(d1.seq):
            if ii1<myTLen1:
                if myTemplate1[ii1]=='N':
                    sL1 += [ZZ,c1.lower()]
                else:
                    sL1 += [zz,c1.lower()]
            else:
                sL1 += [zz,c1]
        workSheet1.write_rich_string(j1,1,*sL1,aa)
        addRichText1(d1.diagramMeExcel(),j1,3,workSheet1)
        for itemn1,itemid1 in enumerate('ftslpBCE'):
            myitem1 = getattr(d1,itemid1)
            workSheet1.write(j1,4+4*itemn1,myitem1.s)
            workSheet1.write(j1,4+4*itemn1+1,myitem1.b)
            workSheet1.write(j1,4+4*itemn1+2,myitem1.e)
            workSheet1.write(j1,4+4*itemn1+3,myitem1.e-myitem1.b)
    j1 += 1
    workSheet1.write(j1+1,0,"ParseRabbit Target",DD)
    workSheet1.write(j1+1,1,os.path.basename(myInputFileName1),dd)
    workSheet1.write(j1+1,2,"ParseRabbit Description",DD)
    workSheet1.write(j1+1,3,NarrativeDescription1,dd)
    workSheet1.write(j1+1,4,"ParseRabbit Command Line",DD)
    workSheet1.write(j1+1,5,'||'.join(sys.argv),dd)
    workSheet1.write(j1+1,6,"ParseRabbit Execution Detail",DD)
    workSheet1.write(j1+1,7,'||'.join(vSysLogInfo1.splitlines()),dd)
    workSheet1.write(j1+1,8,"ParseRabbit TimeStamp",DD)
    workSheet1.write(j1+1,9,vnow,dd)
    workSheet1.write(j1+1,10,"ParseRabbit VSGLog",DD)
    workSheet1.write(j1+1,11,'||'.join(open(vDefaultLogFileName).read().splitlines()),dd)
    workSheet1.write(j1+1,12,"ParseRabbit Python Code",DD)
    workSheet1.write(j1+1,13,'||'.join(open(sys.argv[0]).read().splitlines()[1:]),dd)
    workSheet1.set_row(j1+1,18)

    workBook1.close()

## 2024 Andrew Fire (with impetus and input from Emily Greenwald and Drew Galls)