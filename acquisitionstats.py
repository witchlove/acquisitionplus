#!/usr/bin/python2.7
import getopt, sys, urllib, time, os,  glob,  lxml.etree
from lxml.html import builder as E
from fpdf import FPDF, HTMLMixin
from itertools import  tee,  izip
import os
import re
import formic
import fnmatch
import logging
from xlwt import *


class MyFPDF(FPDF, HTMLMixin):
    pass

class statistics:
    def __init__(self,inputdir, files):
        logging.debug('constructing statistics object')
        self.inputdir = inputdir
        self.files = files
        self.totalNumberOfFiles = 0
        self.totalMultipleBeneficiaries = 0
        self.totalFinancialAdjustment = 0
        self.financialAdjustmentData = []
        self.totalNumberPlacedChildren = 0
        self.placedChildrenData = []
        self.totalYoungJobSeeker = 0
        self.youngJobSeekerData = []
        self.totalBirthAllowance = 0
        self.totalChildMissingRelations = 0
        self.totalLegalGround4= 0
        self.totalHavingForms= 0
        self.havingFormsData=[]
        self.totalReceiver = 0
        self.totalDeceasedFileOwner  = 0
        self.totalVarious = 0
        self.totalAgeAllowance = 0
        self.totalInAssimilation = 0
        self.inAssimilationData=[]
        self.variousData={}
        self.missingFicticiousChildData = []
        self.postfix = "-changes"

    def printHtml(self):
        print "HTML"
        print "Files recieved {0}".format(self.totalNumberOfFiles)
        print "Files having multiple beneficiaries {0}".format(self.totalMultipleBeneficiaries)
        print "Files having FinancialAdjustment(Debt / Withholding) {0}".format(self.totalFinancialAdjustment)
        print "Files having placed children {0}".format(self.totalNumberPlacedChildren)
        print "Files having young jobseeker {0}".format(self.totalYoungJobSeeker)
        print "Files having birth allowance {0}".format(self.totalBirthAllowance)
        print "Files having unemployed child legal ground 4 {0}".format(self.totalLegalGround4)
        print "Files having child(ren) with missing relations {0}".format(self.totalChildMissingRelations)
        print "Files having child(ren) with missing forms {0}".format(self.totalHavingForms)
        print "Files having child(ren) with receiver {0}".format(self.totalReceiver)
        print "Files having deceased fileowner {0}".format(self.totalDeceasedFileOwner)

    def addToList(self, str_to_add, parent, fileName, personINSS):
        if str_to_add in self.variousData:
            self.variousData.get(str_to_add).append((personINSS, fileName, parent))
        else:
            self.variousData[str_to_add] = [(personINSS, fileName, parent)]

    def createPdfreport(self):
        html = """
                <div  align="left" width="90%">directory : {0}</div>
                <table border="1" align="left" width="90%">
                <thead><tr><th width="70%">Check </th><th width="30%">Result</th></tr></thead>
                <tbody>
                <tr><td>Files received </td><td>{1}</td></tr>
                <tr><td>Files having multiple beneficiaries </td><td>{2}</td></tr>
                <tr><td>Files having FinancialAdjustment(Debt / Withholding) </td><td>{3}</td></tr>
                <tr><td>Files having placed children </td><td>{4}</td></tr>
                <tr><td>Files having young jobseeker</td><td>{5}</td></tr>
                <tr><td>Files having birth allowance </td><td>{6}</td></tr>
                <tr><td>Files having unemployed child legal ground 4 </td><td>{7}</td></tr>
                <tr><td>Files having child(ren) with missing relations</td><td>{8}</td></tr>
                <tr><td>Files having child(ren) with forms </td><td>{9}</td></tr>
                <tr><td>Files having child(ren) with receiver </td><td>{10}</td></tr>
                <tr><td>Files having deceased fileowner </td><td>{11}</td></tr>
                <tr><td>Files having various tag </td><td>{12}</td></tr>
                <tr><td>Files having Age Allowance </td><td>{13}</td></tr>
                <tr><td>Files having Child In Assimilation </td><td>{14}</td></tr>
                </tbody>
                </table>
                """.format(self.inputdir, 
                                self.totalNumberOfFiles, 
                                self.totalMultipleBeneficiaries, 
                                self.totalFinancialAdjustment, 
                                self.totalNumberPlacedChildren, 
                                self.totalYoungJobSeeker, 
                                self.totalBirthAllowance,
                                self.totalLegalGround4,
                                self.totalChildMissingRelations,
                                self.totalHavingForms,
                                self.totalReceiver,
                                self.totalDeceasedFileOwner, 
                                self.totalVarious, 
                                self.totalAgeAllowance, 
                                self.totalInAssimilation
                                )
                
        return html
    
    def createVariousDataList(self):
        html = u"""<p>Different value's in various tag</p>"""
        for item in  self.variousData.items():
            html +=u"""<table border="1" align="center" width="100%">"""
            html +=u"""<thead><tr><th width="100%">Various data</th></thead>"""
            html += u"""<tr bgcolor="#FF0000"><td> Data : {0}</td></tr>""".format(item[0])
            for element in item[1]:
                 html += u"""<tr><td>File :{0} </td></tr>""".format(element[1])
                 html += u"""<tr><td>TagLocation : {0}</td></tr>""".format(element[2])
                 html += u"""<tr><td>INSS :{0}</td></tr>""".format(element[0])
            html += u"</table>"
        return html

    def createVariousDataListExcel(self):
        print("not implemented")

    def createMissingFicticiousChildList(self):
        html = u"<p>Files missing ficticious child(ren)</p><ol>"
        for var in self.missingFicticiousChildData:
            html += u"<li>{0}</li>".format(var)
        
        html += u"</ol>"
        return html

    def analyze(self):
        self.findMultipleBeneficiary()
        self.findPlacedChilderen()
        self.findYoungJobSeeker()
        self.findBirthAllowance()
        self.findChildMissingRelations()
        self.findFinancialAdjustment()
        self.findChildLegalGround4()
        self.findFilesHavingForms()
        self.findChildWithReceiver()
        self.findDeceasedFileOwner()
        self.findVarious()
        self.findAgeAllowance()
        self.findChildInAssimilation()
        self.findMissingFicticiousChildren()
    
    
    def hasAcquiPlusChanges(self, filenameForDir, filefilter, xpathToCheck, dirToCheck):
        changeDir = filenameForDir + self.postfix
        change_dir_location = os.path.join(os.path.join(self.inputdir, dirToCheck), changeDir)
        if os.path.exists(change_dir_location):
            logging.debug('checking %s for acquiplus changes', change_dir_location)
            fileset= formic.FileSet(include=filefilter, directory=change_dir_location)
            for filename in fileset.qualified_files(absolute=True):
                doc = lxml.etree.parse(filename)
                for xpathExpression in xpathToCheck:
                    count = doc.xpath(xpathExpression)
                    print count
                    if count > 0 :
                        return 1
                return 0
    
    def checkFormsAddedByAcquiPlus(self, inssFileOwner, filefilter, formCount, childCount):
         changeDir = self.prefix + inssFileOwner + self.postfix
         change_dir_location = os.path.join(self.inputdir, changeDir)
         if os.path.exists(change_dir_location):
            fileset= formic.FileSet(include=filefilter, directory=change_dir_location)
            
    
    def countFiles(self):
        self.totalNumberOfFiles = len(self.files)
        print "total number of xml files recieved {0}".format(self.totalNumberOfFiles)

    def findMultipleBeneficiary(self):
        #print 'finding multiple beneficiaries'
        file_most_bene = ''
        highestcount = 0
        for file in self.files:
            self.totalNumberOfFiles += 1
            current_file = os.path.join(self.inputdir, file)
            doc = lxml.etree.parse(current_file)
            count = doc.xpath('count(//Beneficiary/NaturalPerson)')
            if  count > 1: 
                if highestcount < count:
                    highestcount = count
                    file_most_bene= current_file
                self.totalMultipleBeneficiaries += 1
        print file_most_bene
    
    def findFinancialAdjustment(self):
        #print 'finding financial adjustment':
        for dir,  file_name in self.files.files():
            current_file = os.path.join(os.path.join(self.inputdir, dir), file_name)
            doc = lxml.etree.parse(current_file)
            fileownerINSS = doc.xpath('/FileDescription/FileOwner/PersonINSS/text()')
            if  (doc.xpath('count(//Beneficiary/NaturalPerson/FinancialAdjustment)') > 0 or doc.xpath('count(//Beneficiary/Organization/FinancialAdjustment)') > 0): 
                self.totalFinancialAdjustment += 1
                self.financialAdjustmentData .append((fileownerINSS[0], current_file))
            elif self.hasAcquiPlusChanges(file_name, "*_BeneficiariesFinancialAdjustments.xml",["count(//financialAdjustments)"], dir):
                self.totalFinancialAdjustment += 1
                self.financialAdjustmentData .append((fileownerINSS[0], current_file))
    
    def findPlacedChilderen(self):
        #print 'finding placed children'
        for file in  self.files:
            current_file = os.path.join(self.inputdir, file)
            doc = lxml.etree.parse(current_file)
            count = doc.xpath('count(//PlacedInOrganization)')
            fileownerINSS = doc.xpath('/FileDescription/FileOwner/PersonINSS/text()')
            if  count > 0: 
               self.totalNumberPlacedChildren +=1
               placedChildrenData.append(fileownerINSS[0], current_file)

    def findYoungJobSeeker(self):
        #print 'finding YoungJobSeeker'
        for dir,  file_name in self.files.files():
            current_file = os.path.join(os.path.join(self.inputdir, dir), file_name)
            doc = lxml.etree.parse(current_file)
            fileownerINSS = doc.xpath('/FileDescription/FileOwner/PersonINSS/text()')
            count = doc.xpath('count(//YoungJobSeekerInscriptiondate)')
            if  count >=1 or self.hasAcquiPlusChanges(file_name, "*_YoungJobSeekers.xml", ["count(//inscriptionDate)"],dir):
               self.totalYoungJobSeeker +=1
               self.youngJobSeekerData.append((fileownerINSS[0], current_file))
               
    def findBirthAllowance(self):
        #print 'finding birthallowance'
        for file in self.files:
            current_file = os.path.join(self.inputdir, file)
            doc = lxml.etree.parse(current_file)
            count = doc.xpath('count(//BirthAllowance)')
            if  count >=1: 
               self.totalBirthAllowance +=1
    
    def findChildMissingRelations(self):
        #print 'finding child missing relations'
        for file in self.files:
            current_file = os.path.join(self.inputdir, file)
            doc = lxml.etree.parse(current_file)
            countBene = doc.xpath('count(//BeneficiaryList/Beneficiary)')
            countChild =  doc.xpath('count(//Child)')
            count = doc.xpath('count(//BondBeneficiary/RelationBeneficiarytoChild)')
            if  count != (countBene * countChild): 
               self.totalChildMissingRelations +=1

    def findChildLegalGround4(self):
        #print 'finding child legal ground 4'
        for file in self.files:
            current_file = os.path.join(self.inputdir, file)
            doc = lxml.etree.parse(current_file)
            count = doc.xpath('''count(//Ground[text() = '4'])''')
            if  count > 0  : 
               self.totalLegalGround4 +=1
    
    def findFilesHavingForms(self):
        #print 'finding child missing forms'
        for dir, file_name in self.files.files():
            current_file = os.path.join(os.path.join(self.inputdir, dir), file_name)
            doc = lxml.etree.parse(current_file)
            count = doc.xpath('count(//Child/Forms)')
            fileownerINSS = doc.xpath('/FileDescription/FileOwner/PersonINSS/text()')
            if  count > 0 or self.hasAcquiPlusChanges(file_name, "*_Forms.xml", ["count(//childForms)"],dir):
                self.totalHavingForms +=1
                self.havingFormsData.append((fileownerINSS, current_file))
    
    def findChildWithReceiver(self):
        #print 'finding child with receiver'
        for file in self.files:
            current_file = os.path.join(self.inputdir, file)
            doc = lxml.etree.parse(current_file)
            count = doc.xpath('count(//Child/ReceiverTypeList)')
            if  count > 0  : 
               self.totalReceiver +=1
    
    def findDeceasedFileOwner(self):
        #print 'finding child with deceased fileowner'
        for file in self.files:
            current_file = os.path.join(self.inputdir, file)
            doc = lxml.etree.parse(current_file)
            count = doc.xpath('count(//FileOwner/PersonDateOfDeath)')
            if  count > 0  : 
               self.totalDeceasedFileOwner +=1
    
    def findVarious(self):
        #print 'finding files with  various'
        for file in self.files:
            current_file = os.path.join(self.inputdir, file)
            doc = lxml.etree.parse(current_file)
            count = doc.xpath('count(//Various)')
            if  count > 0  : 
                fileownerINSS = doc.xpath('/FileDescription/FileOwner/PersonINSS/text()')
                variousElementList = doc.xpath('//Various')
                for el in variousElementList:
                    self.addToList(el.text, self.createPath(el, []), file,  fileownerINSS[0])
                    self.totalVarious +=1
    
    def createPath(self, el, fullPath):
        if(el.getparent() is not None):
            fullPath.append(el.getparent().tag)
            return self.createPath(el.getparent(), fullPath)
        else:
            path = ""
            for el in reversed(fullPath):
                path += el + "->"
            return path[:-2]
    
    def findAgeAllowance(self):
        #print 'finding files with  AgeAllowance'
        for file in self.files:
            current_file = os.path.join(self.inputdir, file)
            doc = lxml.etree.parse(current_file)
            count = doc.xpath('count(//Child/AgeAllowance)')
            if  count > 0  : 
               self.totalAgeAllowance +=1
    
    def findChildInAssimilation(self):
        for file in self.files:
            current_file = os.path.join(self.inputdir, file)
            doc = lxml.etree.parse(current_file)
            fileownerINSS = doc.xpath('/FileDescription/FileOwner/PersonINSS/text()')
            count = doc.xpath('count(//Child/ChildInAssimilation)')
            if  count > 0  : 
               self.totalInAssimilation +=1
               self.inAssimilationData.append((fileownerINSS[0], current_file))
    
    def findMissingFicticiousChildren(self):
         for file in self.files:
            current_file = os.path.join(self.inputdir, file)
            doc = lxml.etree.parse(current_file)
            payedRang = doc.xpath('//ChildList/Child/Rang/PayedRangChild/text()')
            fileownerINSS = doc.xpath('/FileDescription/FileOwner/PersonINSS/text()')
            if  payedRang :
                if  not '1' in payedRang :
                    self.missingFicticiousChildData.append((fileownerINSS, current_file))
                elif  self.checkSequence(payedRang) == 0:
                    self.missingFicticiousChildData.append((fileownerINSS, current_file))
    
    def pairwise(self, iterable):
        "s -> (s0,s1), (s1,s2), (s2, s3), ..."
        a, b = tee(iterable)
        next(b, None)
        return izip(a, b)
    
    def checkSequence(self,  range):
        sortedList = sorted(set(range))
        for x, y in self.pairwise(sortedList) :
            if int(y)-int(x) != 1 :
                return 0
        return 1
        
    def writeExcelFile(self):
        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 4
        fnt.bold = True
        
        borders = Borders()
        borders.left = 6
        borders.right = 6
        borders.top = 6
        borders.bottom = 6
        
        al = Alignment()
        al.horz = Alignment.HORZ_CENTER
        al.vert = Alignment.VERT_CENTER
        
        style = XFStyle()
        style.font = fnt
        style.borders = borders
        style.alignment = al
        wb = Workbook()
        ws = wb.add_sheet("Various Data")
        ws2 = wb.add_sheet("Files Youngjobseeker")
        ws3 = wb.add_sheet("Files Child In Assimilation")
        ws4 =  wb.add_sheet("Files Child In Placement")
        ws5 = wb.add_sheet("Financial adjustment")
        ws6 = wb.add_sheet("Files with Forms")
        ws7 = wb.add_sheet("Files missing fictitious")
        rowCounter=0
        for item in self.variousData.items() :
            ws.write_merge(rowCounter, rowCounter, 0, 2, item[0], style)
            rowCounter +=1
            for element in item[1]:
                colCounter=0
                for colVal in element:
                    ws.write(rowCounter, colCounter, colVal)
                    colCounter += 1
                rowCounter += 1
        ws.col(0).width = 256 * 5
    
        ws2.write(0, 0, "INSS FILEOWNER")
        ws2.write(0, 1, "FILE")
        ws2.col(0).width = 256 * len("INSS FILEOWNER")
        ws2.col(1).width = 256 * max([len(row[1]) for row in self.youngJobSeekerData])
        rowCounter=1
        for item in self.youngJobSeekerData:
            colCounter =0
            for colVal in item :
                ws2.write(rowCounter,colCounter, colVal )
                colCounter += 1
            rowCounter += 1 
        """write excal tab children in assimilation"""
        
        ws3.write(0, 0, "INSS FILEOWNER")
        ws3.write(0, 1, "FILE")
        ws3.col(0).width = 256 * len("INSS FILEOWNER")
        ws3.col(1).width = 256 * max([len(row[1]) for row in self.inAssimilationData])
        rowCounter=1
        for item in self.inAssimilationData:
            colCounter =0
            for colVal in item :
                ws3.write(rowCounter,colCounter, colVal )
                colCounter += 1
            rowCounter += 1
        
        ws4.write(0, 0, "INSS FILEOWNER")
        ws4.write(0, 1, "FILE")
        ws4.col(0).width = 256 * len("INSS FILEOWNER")
        if len(self.placedChildrenData) > 0:
            ws4.col(1).width = 256 * max([len(row[1]) for row in self.placedChildrenData])
        rowCounter=1
        for item in self.placedChildrenData:
            colCounter =0
            for colVal in item :
                ws4.write(rowCounter,colCounter, colVal )
                colCounter += 1
            rowCounter += 1
        
        ws5.write(0, 0, "INSS FILEOWNER")
        ws5.write(0, 1, "FILE")
        ws5.col(0).width = 256 * len("INSS FILEOWNER")
        if len(self.financialAdjustmentData) > 0:
            ws5.col(1).width = 256 * max([len(row[1]) for row in self.financialAdjustmentData])
        rowCounter=1
        for item in self.financialAdjustmentData:
            colCounter =0
            for colVal in item :
                ws5.write(rowCounter,colCounter, colVal )
                colCounter += 1
            rowCounter += 1
        
        ws6.write(0, 0, "INSS FILEOWNER")
        ws6.write(0, 1, "FILE")
        ws6.col(0).width = 256 * len("INSS FILEOWNER")
        if len(self.havingFormsData) > 0:
            ws6.col(1).width = 256 * max([len(row[1]) for row in self.havingFormsData])
        rowCounter=1
        for item in self.havingFormsData:
            colCounter =0
            for colVal in item :
                ws6.write(rowCounter,colCounter, colVal )
                colCounter += 1
            rowCounter += 1
            
        ws7.write(0, 0, "INSS FILEOWNER")
        ws7.write(0, 1, "FILE")
        ws7.col(0).width = 256 * len("INSS FILEOWNER")
        if len(self.missingFicticiousChildData) > 0:
            ws7.col(1).width = 256 * max([len(row[1]) for row in self.missingFicticiousChildData])
        rowCounter=1
        for item in self.missingFicticiousChildData:
            colCounter =0
            for colVal in item :
                ws7.write(rowCounter,colCounter, colVal )
                colCounter += 1
            rowCounter += 1
        
        
        wb.save("myworkbook.xls")
    
def main(argv):
    
    recursive = 0
   
    logging.basicConfig(filename='stats.log',level=logging.DEBUG)
    logging.info('Starting stats script')

    try:
      opts, args = getopt.getopt(argv,"VhRi:o:",["idir=","odir="])
      
    except getopt.GetoptError:
            print 'test.py -i <inputdir>'
            sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print 'test.py -i <inputdir> -o <outputdir>'
            sys.exit()
        elif opt in ("-i", "--idir"):
            inputdir= arg
        elif opt in ("-o", "--odir"):
            outputdir = arg
        elif opt in ("-R"):
            recursive=1
        elif opt in ("-V"):
            verbose=1
            
    logging.debug('Input file is %s', inputdir)
   
    htmlparts = []
    
    if recursive == 1:
        "use recursive to get all reporting"
        stats = statistics(inputdir, formic.FileSet(include="*.xml",exclude=["*.xml-changes","*_*_*_*.xml" ], directory=inputdir))
        stats.analyze()
        stats. writeExcelFile()
        htmlparts.append(stats.createPdfreport())
        #htmlparts.append(stats.createVariousDataList())
        htmlparts.append(stats.createMissingFicticiousChildList())
    else:
        for root, dirs, files in os.walk(inputdir, topdown=True):
            for name in dirs:
                fullPath = os.path.join(root, name)
                filesInDir = os.listdir(fullPath)
                filteredFiles = []
                for f in filesInDir:
                    if(f.endswith(".xml")):
                        filteredFiles.append(f)
                stats = statistics(fullPath, filteredFiles)
                stats.analyze()
                htmlparts.append(stats.createPdfreport())
    
    pdf=MyFPDF()
    pdf.add_page()
    
    pdf.write_html('''<H1 align="center">Statistics Report</H1>''')
    for html in htmlparts:
        pdf.write_html(html)
    
    pdf.write
    
    reportLocation = os.path.join(outputdir, "statisticsreport.pdf")
    pdf.output(reportLocation,'F')
    
    print "Done"
    
if __name__ == "__main__":
    main(sys.argv[1:])
