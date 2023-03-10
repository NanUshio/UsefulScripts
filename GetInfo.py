# Example:
# python GetInfo.py --EncoderName VVENC 
#                   --ReadFilePath ./RA_AllFrame_AllClass_MCTF_ON_ANCHOR/logs/ 
#                   --ReadFileType .txt 
#                   --WriteFilePath ./RA_AllFrame_AllClass_MCTF_ON_ANCHOR/ 
#                   --WriteFileName result 
#                   --WriteFileType .log
#                   --LogLevel      1
#


import argparse
import os
import re
import fnmatch
import pandas
import openpyxl
from openpyxl import Workbook , load_workbook

vvcvideoDict =  ["Tango2"             ,
                 "FoodMarket4"        ,
                 "Campfire"           ,
                 "CatRobot"           ,
                 "DaylightRoad2"      ,
                 "ParkRunning3"       ,
                 "MarketPlace"        ,
                 "RitualDance"        ,
                 "Cactus"             ,
                 "BasketballDrive"    ,
                 "BQTerrace"          ,
                 "RaceHorseC"        ,
                 "BQMall"             ,
                 "PartyScene"         ,
                 "BasketballDrill"    ,
                 "RaceHorses"         ,
                 "BQSquare"           ,
                 "BlowingBubbles"     ,
                 "BasketballPass"     ,
                 "FourPeople"         ,
                 "Johnny"             ,
                 "KristenAndSara"     ,
                 "ArenaOfValor"       ,
                 "BasketballDrillText",
                 "SlideEditing"       ,
                 "SlideShow"]

hevcvideoDict =  ["Traffic"            ,
                  "PeopleOnStreet"     ,
#                  "Nebuta"             ,
#                  "SteamLocomotive"    ,
                  "Kimono"             ,
                  "ParkScene"          ,
                  "Cactus"             ,
                  "BasketballDrive"    ,
                  "BQTerrace"          ,
                  "BasketballDrill"    ,
                  "BQMall"             ,
                  "PartyScene"         ,
                  "RaceHorseC"        ,
                  "BasketballPass"     ,
                  "BQSquare"           ,
                  "BlowingBubbles"     ,
                  "RaceHorses"         ,
                  "FourPeople"         ,
                  "Johnny"             ,
                  "KristenAndSara"     ,
                  "BasketballDrillText",
                  "ChinaSpeed"         ,
                  "SlideEditing"       ,
                  "SlideShow"]       

def parse_args() :
    '''
    Parsing Command-Line Arguments
    :return args : Dict < ArgName : ArgValue >
    '''
    parser = argparse.ArgumentParser()
    # require target encoder log path
    parser.add_argument('--ReadFilePath', 
                        type=str, 
                        default='./', 
                        help='Path of Encoder Files for GetInfo.py to Read',
                        required=True
                        )
    parser.add_argument('--ReadFileType', 
                        type=str, 
                        default='log', 
                        help='Type of Encoder Files for GetInfo.py to Read',
                        required=True
                        )    
    parser.add_argument('--WriteFilePath',
                        type=str,
                        default='./',
                        help='Path of GetInfo.py to Write Result File',
                        required=False
                        )
    parser.add_argument('--WriteFileType', 
                        type=str, 
                        default='log', 
                        help='Type of Encoder Files for GetInfo.py to Write',
                        required=True
                        )
    parser.add_argument('--WriteFileName',
                        type=str,
                        default='result',
                        help='Name of Result File of GetInfo.py',
                        required=False,
                        )
    parser.add_argument('--LogLevel',
                        type=int,
                        default='1',
                        help='Level of GetInfo.py to Read Information',
                        required=False,
                        )
    parser.add_argument('--EncoderName',
                        type=str,
                        default='HM',
                        help='Name of Encoder for GetInfo.py to Read',
                        required=True,
                        )
    parser.add_argument('--CtcType',
                        type=str,
                        default='HEVC',
                        help='CTC Type Used',
                        required=False,
                        )
    args = parser.parse_args()
    return args

def traverse_files( TargetFilePath , TargetFileType ) :
    '''
    Traverse Files with TargetFileType under TargetFilePath
    :return targetFileList : [ TargetFilePath/FileName1.TargetFileType , ... , TargetFilePath/FileNameN.TargetFileType]
    ''' 
    targetFileNameList = []
    for root , dics , fils in os.walk( TargetFilePath ) :
        for filename in fils :
            name , suf = os.path.splitext( filename )
            if suf == TargetFileType :
                targetFileNameList.append( os.path.join( root, filename ) )
    return targetFileNameList

class EncInfo:
    '''
    Class to Store Encoder's Ouput Information Read from TargetFileList
    ''' 
    def __init__(self) :
        self.SeqName       = ""
        self.SeqAvgQp      = ""
        self.AvgBitRate    = ""
        self.AvgYUVPsnr    = ""
        self.AvgYPsnr      = ""
        self.AvgUPsnr      = ""
        self.AvgVPsnr      = ""
        self.EncTime       = ""
        self.FrameTypeList = []
        self.QPList        = []
        self.BitRateList   = []      
        self.YPsnrList     = []
        self.UPsnrList     = []
        self.VPsnrList     = []
        self.YUVPsnrList   = []
        self.GradientList  = []
        self.AvgGradient   = ""
    def __non_zero__(self) :
        return bool ( 
                    self.SeqName       or
                    self.AvgBitRate    or
                    self.AvgYUVPsnr    or
                    self.AvgYPsnr      or
                    self.AvgUPsnr      or
                    self.AvgVPsnr      or
                    self.EncTime       or
                    self.FrameTypeList or
                    self.YPsnrList     or
                    self.UPsnrList     or
                    self.VPsnrList     or
                    self.YUVPsnrList   or
                    self.BitRateList   or
                    self.QPList        or
                    self.GradientList  or
                    self.AvgGradient
                    )

def search_result(chunk, matchSyntax, index):
    '''
    Find MatchSyntax in Chunk and Read matchWord by index
    :return matchWord
    ''' 
    if re.search(matchSyntax, chunk):
        splitWord = re.split(r'\s+', chunk)
        matchWord = []
        matchWord.append(splitWord[index])
        return matchWord

def read_information_vvenc( TargetFile , LogLevel ) :
    '''
    Collect VVENC Encoder's Output Information in TargetFile
    :return encInfo
    ''' 
    encInfo = EncInfo()
    if os.path.isfile( TargetFile ):
        findResult = re.match( r'.*\/(\D+_-?\d+).*\.txt', TargetFile )
        if findResult :
            encInfo.SeqName = findResult.group(1)
        findResult = re.match( r'.*\/(\D+_\d+).*\.log', TargetFile )
        if findResult :
            encInfo.SeqName = findResult.group(1)
        findResult = re.match( r'.*\/(\D+_\d+).*\.csv', TargetFile )
        if findResult :
            encInfo.SeqName = findResult.group(1)
        findResult = re.match( r'.*_(\d+)',encInfo.SeqName)
        if findResult :
            encInfo.SeqAvgQp = findResult.group(1)
        fileHandle = open( TargetFile , 'r' )
        while 1:
            chunk = fileHandle.readline()
            if chunk == '' :  break
            if LogLevel > 0 :
                matchWord  = '\sFrames'
                findResult = search_result(chunk, matchWord, 0)
                if findResult:
                    chunk  = fileHandle.readline()
                    result = re.split(r'\s+', chunk)
                    encInfo.AvgBitRate = result[3]
                    encInfo.AvgYPsnr   = result[4]
                    encInfo.AvgUPsnr   = result[5]
                    encInfo.AvgVPsnr   = result[6]
                    encInfo.AvgYUVPsnr = result[7]
                matchWord  = '\sTime'
                findResult = search_result(chunk, matchWord, 3)
                if findResult:
                    encInfo.EncTime = findResult[0]
            if LogLevel > 1 :
                matchWord  = 'POC'
                findResult = search_result(chunk, matchWord, 0)
                if findResult :
                    result = re.split(r'\s+',chunk)
                    for index in range( len( result ) ) :
                        if re.match( r'.*SLICE', result[index] ) :
                            encInfo.FrameTypeList.append( result[index][:-1] )
                        elif result[index] == 'QP' :
                            encInfo.QPList.append( result[index + 1][:-1] )
                        elif result[index] == 'bits' :
                            encInfo.BitRateList.append( result[index - 1] )
                        elif result[index] == '[Y' :
                            encInfo.YPsnrList.append( result[index + 1] )
                        elif result[index] == 'U' :
                            encInfo.UPsnrList.append( result[index + 1] )
                        elif result[index] == 'V' :
                            encInfo.VPsnrList.append( result[index + 1] )
                    encInfo.YUVPsnrList.append( ( 6 * float(encInfo.YPsnrList[-1]) + float(encInfo.UPsnrList[-1]) + float(encInfo.VPsnrList[-1]) ) / 8 )
                    #if encInfo.FrameTypeList[-1] == 'I-SLICE' :
                    #    del(encInfo.FrameTypeList[-1])
                    #    del(encInfo.QPList[-1]) 
                    #    del(encInfo.BitRateList[-1]) 
                    #    del(encInfo.YPsnrList[-1]) 
                    #    del(encInfo.UPsnrList[-1]) 
                    #    del(encInfo.VPsnrList[-1]) 
                    #    del(encInfo.YUVPsnrList[-1])
    assert encInfo , 'Empty encInfo T T'
    return encInfo

def read_information_vtm( TargetFile , LogLevel ) :
    '''
    Collect VTM Encoder's Output Information in TargetFile
    :return encInfo
    ''' 
    encInfo = EncInfo()
    if os.path.isfile( TargetFile ):
        findResult = re.match( r'.*\/(\D+_\d+)\.txt', TargetFile )
        if findResult :
            encInfo.SeqName = findResult.group(1)
        findResult = re.match( r'.*\/(.*_\d+)\.log', TargetFile )
        if findResult :
            encInfo.SeqName = findResult.group(1)
        findResult = re.match( r'.*\/(\D+_\d+)\.csv', TargetFile )
        if findResult :
            encInfo.SeqName = findResult.group(1)
        findResult = re.match( r'.*_(\d+)',encInfo.SeqName)
        if findResult :
            encInfo.SeqAvgQp = findResult.group(1)
        fileHandle = open( TargetFile , 'r' )
        while 1:
            chunk = fileHandle.readline()
            if chunk == '' :  break
            if LogLevel > 0 :
                matchWord  = '\sFrames'
                findResult = search_result(chunk, matchWord, 0)
                if findResult:
                    chunk  = fileHandle.readline()
                    result = re.split(r'\s+', chunk)
                    encInfo.AvgBitRate = result[3]
                    encInfo.AvgYPsnr   = result[4]
                    encInfo.AvgUPsnr   = result[5]
                    encInfo.AvgVPsnr   = result[6]
                    encInfo.AvgYUVPsnr = result[7]
                matchWord  = '\sTime'
                findResult = search_result(chunk, matchWord, 3)
                if findResult:
                    encInfo.EncTime = findResult[0]
            if LogLevel > 1 :
                matchWord  = 'Gradient'
                findResult = search_result(chunk, matchWord, 0)
                if findResult :
                    result = re.split(r'\s+',chunk)
                    for index in range( len( result ) ) :
                        if re.match( r'.*SLICE', result[index] ) :
                            encInfo.FrameTypeList.append( result[index][:-1] )
                        elif result[index] == 'QP' :
                            encInfo.QPList.append( result[index + 1])
                        elif result[index] == 'bits' :
                            encInfo.BitRateList.append( result[index - 1] )
                        elif result[index] == '[Y' :
                            encInfo.YPsnrList.append( result[index + 1] )
                        elif result[index] == 'U' :
                            encInfo.UPsnrList.append( result[index + 1] )
                        elif result[index] == 'V' :   
                            encInfo.VPsnrList.append( result[index + 1] )
                        elif result[index] == 'Gradient' :
                            encInfo.GradientList.append( result[index + 1] )
                        elif result[index] == 'Avg' :
                            encInfo.AvgGradient = result[index + 1]
                    encInfo.YUVPsnrList.append( ( 6 * float(encInfo.YPsnrList[-1]) + float(encInfo.UPsnrList[-1]) + float(encInfo.VPsnrList[-1]) ) / 8 )  
    assert encInfo , 'Empty encInfo T T'
    return encInfo

def read_information_hm( TargetFile , LogLevel ) :
    '''
    Collect HM Encoder's Output Information in TargetFile
    :return encInfo
    ''' 
    encInfo = EncInfo()
    if os.path.isfile( TargetFile ):
        findResult = re.match( r'.*\/(.+?)\.txt', TargetFile )
        if findResult :
            encInfo.SeqName = findResult.group(1)
        findResult = re.match( r'.*\/(.*_\d+)\.log', TargetFile )
        if findResult :
            encInfo.SeqName = findResult.group(1)
        findResult = re.match( r'.*\/(\D+_\d+).*\.csv', TargetFile )
        if findResult :
            encInfo.SeqName = findResult.group(1)
        findResult = re.match( r'.*_(\d+)',encInfo.SeqName)
        if findResult :
            encInfo.SeqAvgQp = findResult.group(1)
        fileHandle = open( TargetFile , 'r' )
        while 1:
            chunk = fileHandle.readline()
            if chunk == '' :  break
            if LogLevel > 0 :
                matchWord  = '\sFrames'
                findResult = search_result(chunk, matchWord, 0)
                if findResult:
                    chunk  = fileHandle.readline()
                    result = re.split(r'\s+', chunk)
                    if result[2] == 'a' :
                        encInfo.AvgBitRate = result[3]
                        encInfo.AvgYPsnr   = result[4]
                        encInfo.AvgUPsnr   = result[5]
                        encInfo.AvgVPsnr   = result[6]
                        encInfo.AvgYUVPsnr = result[7]
                matchWord  = '\sTime'
                findResult = search_result(chunk, matchWord, 3)
                if findResult:
                    encInfo.EncTime = findResult[0]
            if LogLevel > 1 :
                matchWord  = 'Gradient'
                findResult = search_result(chunk, matchWord, 0)
                if findResult :
                    result = re.split(r'\s+',chunk)
                    for index in range( len( result ) ) :
                        if re.match( r'.*SLICE', result[index] ) :
                            encInfo.FrameTypeList.append( result[index][:-1] )
                        elif result[index] == 'QP' :
                            encInfo.QPList.append( result[index + 1])
                        elif result[index] == 'bits' :
                            encInfo.BitRateList.append( result[index - 1] )
                        elif result[index] == '[Y' :
                            encInfo.YPsnrList.append( result[index + 1] )
                        elif result[index] == 'U' :
                            encInfo.UPsnrList.append( result[index + 1] )
                        elif result[index] == 'V' :
                            encInfo.VPsnrList.append( result[index + 1] )
                        elif result[index] == 'Gradient' :
                            encInfo.GradientList.append( result[index + 1] )
                        elif result[index] == 'Avg' :
                            encInfo.AvgGradient = result[index + 1]
                    encInfo.YUVPsnrList.append( ( 6 * float(encInfo.YPsnrList[-1]) + float(encInfo.UPsnrList[-1]) + float(encInfo.VPsnrList[-1]) ) / 8 ) 
    assert encInfo , 'Empty encInfo T T'
    return encInfo

def read_information_x265( TargetFile , LogLevel ) :
    '''
    Collect X265 Encoder's Output Information in TargetFile
    :return encInfo
    ''' 
    encInfo = EncInfo()
    if os.path.isfile( TargetFile ):
        findResult = re.match( r'.*\/(.+?)\.txt', TargetFile )
        if findResult :
            encInfo.SeqName = findResult.group(1)
        findResult = re.match( r'.*\/(\D+_\d+).*\.log', TargetFile )
        if findResult :
            encInfo.SeqName = findResult.group(1)
        findResult = re.match( r'.*\/(\D+_\d+).*\.csv', TargetFile )
        if findResult :
            encInfo.SeqName = findResult.group(1)
        findResult = re.match( r'.*_(\d+)',encInfo.SeqName)
        if findResult :
            encInfo.SeqAvgQp = findResult.group(1)
        fileHandle = open( TargetFile , 'r' )
        chunks = fileHandle.readlines()
        for index in range( len(chunks) ) :
            if chunks[index] == '\n' :
                continue
            if LogLevel > 0 :
                if index == len(chunks) - 1 :
                    result = re.split(r',\s*', chunks[index])
                    encInfo.AvgBitRate = result[4]
                    encInfo.AvgYPsnr   = result[5]
                    encInfo.AvgUPsnr   = result[6]
                    encInfo.AvgVPsnr   = result[7]
                    encInfo.AvgYUVPsnr = result[8]
                    encInfo.EncTime    = result[2]
                    continue
            if LogLevel > 1 :
                matchWord  = 'Encode'
                findResult = search_result(chunks[index], matchWord, 0)
                if findResult :
                    continue
                matchWord = 'Summary'
                findResult = search_result(chunks[index], matchWord, 0)
                if findResult :
                    continue
                matchWord = 'Command'
                findResult = search_result(chunks[index], matchWord, 0)
                if findResult :
                    continue
                result = re.split(r',\s*',chunks[index])
                encInfo.FrameTypeList.append( result[1] )
                encInfo.QPList.append( result[3])
                encInfo.BitRateList.append( result[4] )
                encInfo.YPsnrList.append( result[7] )
                encInfo.UPsnrList.append( result[8] )
                encInfo.VPsnrList.append( result[9] )
                encInfo.YUVPsnrList.append( ( 6 * float(encInfo.YPsnrList[-1]) + float(encInfo.UPsnrList[-1]) + float(encInfo.VPsnrList[-1]) ) / 8 ) 
    assert encInfo , 'Empty encInfo T T'
    return encInfo

read_information_enctype = {
    'HM'    : read_information_hm    ,
    'VTM'   : read_information_vtm   ,
    'VVENC' : read_information_vvenc ,
    'X265'  : read_information_x265  ,
}

def read_information( TargetFileList , EncType , LogLevel ) :
    '''
    Read Information from Files in TargetFileList
    :return encInfoList : [ EncInfo1 , ... , EncInfoN ]
    ''' 
    encInfoList = []
    for targetFile in TargetFileList :
        encInfo = read_information_enctype.get( EncType )( targetFile , LogLevel )
        encInfoList.append( encInfo )
    return encInfoList

def delete_nonsequence( EncInfoList ) :
    '''
    Delete encInfo with empty SeqName in EncInfoList
    :return encInfoList : [ EncInfo1 , ... , EncInfoN ]
    ''' 
    encInfoList = []
    for encInfo in EncInfoList :
        if encInfo.SeqName != '' :
            encInfoList.append( encInfo )
    return encInfoList

def sort_sequence_vvc ( EncInfoList ) :
    '''
    Sort encInfo with SeqName in EncInfoList due to VVC SDR CTC
    :return encInfoList : [ EncInfo1 , ... , EncInfoN ]
    ''' 
    encInfoList = []
    for videoName in vvcvideoDict : 
        regex = re.compile(r'(%s)' %videoName)
        flag  = 1
        for encInfo in EncInfoList : 
            if regex.match( encInfo.SeqName ) :
                encInfoList.append( encInfo )
                flag = 0
        if flag :
            uncodedEncInfo = EncInfo()
            uncodedEncInfo.SeqName = videoName
            encInfoList.append( uncodedEncInfo )
    return encInfoList

def sort_sequence_hevc ( EncInfoList ) :
    '''
    Sort encInfo with SeqName in EncInfoList due to HEVC SDR CTC
    :return encInfoList : [ EncInfo1 , ... , EncInfoN ]
    ''' 
    encInfoList = []
    for encInfo in EncInfoList : 
        for videoName in hevcvideoDict : 
            regex = re.compile(r'(%s)' %videoName)
            if regex.match( encInfo.SeqName ) :
                encInfoList.append( encInfo )
    return encInfoList

sort_sequence_ctc = {
    'HEVC'  : sort_sequence_hevc  ,
    'VVC'   : sort_sequence_vvc   ,
}

def create_file_log( WriteFilePath , WriteFileName ) :
    fullFilePath = WriteFilePath + WriteFileName + '.log'
    if os.path.exists( fullFilePath ) :
        os.remove( fullFilePath )
        os.mknod( fullFilePath )
    elif os.path.exists( WriteFilePath ) :
        os.mknod( fullFilePath )
    else :
        os.mkdir( WriteFilePath )
        os.mknod( fullFilePath )

def create_file_txt( WriteFilePath , WriteFileName ) :
    fullFilePath = WriteFilePath + WriteFileName + '.txt'
    if os.path.exists( fullFilePath ) :
        os.remove( fullFilePath )
        os.mknod( fullFilePath )
    elif os.path.exists( WriteFilePath ) :
        os.mknod( fullFilePath )
    else :
        os.mkdir( WriteFilePath )
        os.mknod( fullFilePath )

def create_file_csv( WriteFilePath , WriteFileName ) :
    None

def create_file_xlsx( WriteFilePath , WriteFileName ) :
    fullFilePath = WriteFilePath + WriteFileName + '.xlsx'
    if os.path.exists( fullFilePath ) :
        os.remove( fullFilePath )
        dataFrame = pandas.DataFrame()
        dataFrame.to_excel( fullFilePath )
    elif os.path.exists( WriteFilePath ) :
        dataFrame = pandas.DataFrame()
        dataFrame.to_excel( fullFilePath )
    else :
        os.mkdir( WriteFilePath )
        dataFrame = pandas.DataFrame()
        dataFrame.to_excel( fullFilePath )


create_file_filetype = {
    '.log'  : create_file_log  ,
    '.txt'  : create_file_txt  ,
    '.csv'  : create_file_csv  ,
    '.xlsx' : create_file_xlsx ,
}

def create_file( WriteFilePath , WriteFileName , WriteFileType ) :
    fullFilePath = WriteFilePath + WriteFileName + WriteFileType
    create_file_filetype.get(WriteFileType)( WriteFilePath , WriteFileName )
    assert os.path.exists( fullFilePath ) , "Target Write File UnCreated T T"

def write_information_log( EncInfoList , WriteFilePath , WriteFileName , LogLevel ) :
   fullFilePath = WriteFilePath + WriteFileName + '.log'
   fileHandle = open( fullFilePath , 'a' )
       
   if LogLevel > 0 :
        # write Header Information
        fileHandle.write( "%-7s \t%-13s \t%-7s \t%-7s \t%-7s \t%-7s \t%-13s \t%-13s \n" %( "Qp" , "BitRate" , "AvgPsnr" , "YPsnr" , "UPsnr" , "VPsnr", "EncTime", "Sequence" ) )
        for encInfo in EncInfoList :
            fileHandle.write( "%-7s \t%-13s \t%-7s \t%-7s \t%-7s \t%-7s \t%-13s \t%-13s \n" %( encInfo.SeqAvgQp , encInfo.AvgBitRate , encInfo.AvgYUVPsnr , encInfo.AvgYPsnr , encInfo.AvgUPsnr , encInfo.AvgVPsnr , encInfo.EncTime , encInfo.SeqName ) )
   if LogLevel > 1 :
        for encInfo in EncInfoList :
            # write Header Information
            fileHandle.write( "%-13s \t%-13s \n" %( "Sequence" , encInfo.SeqName ) )
            fileHandle.write( "%-13s \t%-5s \t%-13s \t%-7s \t%-7s \t%-7s \t%-7s \n" %( "SliceType" , "QP" , "BitRate" , "AvgPsnr" , "YPsnr" , "UPsnr" , "VPsnr" ) )
            for index in range( len( encInfo.QPList ) ) :    
                fileHandle.write( "%-13s \t%-5s \t%-13s \t%-.4f \t%-7s \t%-7s \t%-7s \n" %( encInfo.FrameTypeList[index] , encInfo.QPList[index] , encInfo.BitRateList[index] , \
                                                                                        encInfo.YUVPsnrList[index] , \
                                                                                        encInfo.YPsnrList[index] , encInfo.UPsnrList[index] , encInfo.VPsnrList[index] ) )
   fileHandle.close()
    
def write_information_txt( EncInfoList , WriteFilePath , WriteFileName , LogLevel ) :
   fullFilePath = WriteFilePath + WriteFileName + '.txt'
   fileHandle = open( fullFilePath , 'a' )
       
   if LogLevel > 0 :
        # write Header Information
        fileHandle.write( "%-7s \t%-13s \t%-7s \t%-7s \t%-7s \t%-7s \t%-13s \t%-13s \n" %( "Qp" , "BitRate" , "AvgPsnr" , "YPsnr" , "UPsnr" , "VPsnr", "EncTime", "Sequence" ) )
        for encInfo in EncInfoList :
            fileHandle.write( "%-7s \t%-13s \t%-7s \t%-7s \t%-7s \t%-7s \t%-13s \t%-13s \n" %( encInfo.SeqAvgQp , encInfo.AvgBitRate , encInfo.AvgYUVPsnr , encInfo.AvgYPsnr , encInfo.AvgUPsnr , encInfo.AvgVPsnr , encInfo.EncTime , encInfo.SeqName ) )
   if LogLevel > 1 :
        for encInfo in EncInfoList :
            # write Header Information
            fileHandle.write( "%-13s \t%-13s \n" %( "Sequence" , encInfo.SeqName ) )
            fileHandle.write( "%-13s \t%-5s \t%-13s \t%-7s \t%-7s \t%-7s \t%-7s \n" %( "SliceType" , "QP" , "BitRate" , "AvgPsnr" , "YPsnr" , "UPsnr" , "VPsnr" ) )
            for index in range( len( encInfo.QPList ) ) :    
                fileHandle.write( "%-13s \t%-5s \t%-13s \t%-.4f \t%-7s \t%-7s \t%-7s \n" %( encInfo.FrameTypeList[index] , encInfo.QPList[index] , encInfo.BitRateList[index] , \
                                                                                        encInfo.YUVPsnrList[index] , \
                                                                                        encInfo.YPsnrList[index] , encInfo.UPsnrList[index] , encInfo.VPsnrList[index] ) )
   fileHandle.close()

def write_information_csv( EncInfoList , WriteFilePath , WriteFileName , LogLevel ) :
    None
    
def write_information_xlsx( EncInfoList , WriteFilePath , WriteFileName , LogLevel ) :
    fullFilePath = WriteFilePath + WriteFileName + '.xlsx'
    if LogLevel > 0 :
        # write Header Information
        fileHandle = open( fullFilePath , "a" )
        colHeaderList = [ "Qp" , "BitRate" , "AvgPsnr" , "YPsnr" , "UPsnr" , "VPsnr", "EncTime", "Sequence" , "AvgGradient" ]
        avgBitRateList = []
        avgYUVPsnrList = []
        avgYPsnrList   = []
        avgUPsnrList   = []
        avgVPsnrList   = []
        avgEncTimeList = []
        SeqNameList = []
        SeqAvgQpList = []
        avgGradientList = []
        for encInfo in EncInfoList :
            avgBitRateList.append(encInfo.AvgBitRate) 
            avgYUVPsnrList.append(encInfo.AvgYUVPsnr)
            avgYPsnrList.append(encInfo.AvgYPsnr)  
            avgUPsnrList.append(encInfo.AvgUPsnr)
            avgVPsnrList.append(encInfo.AvgVPsnr)
            avgEncTimeList.append(encInfo.EncTime)
            SeqNameList.append(encInfo.SeqName)
            SeqAvgQpList.append(encInfo.SeqAvgQp)
            avgGradientList.append(encInfo.AvgGradient)
        dataFrame = pandas.DataFrame({ colHeaderList[0] : SeqAvgQpList   ,
                                       colHeaderList[1] : avgBitRateList ,
                                       colHeaderList[2] : avgYUVPsnrList ,
                                       colHeaderList[3] : avgYPsnrList   ,
                                       colHeaderList[4] : avgUPsnrList   ,
                                       colHeaderList[5] : avgVPsnrList   ,
                                       colHeaderList[6] : avgEncTimeList ,
                                       colHeaderList[7] : SeqNameList    ,
                                       colHeaderList[8] : avgGradientList})
        dataFrame.to_excel(fullFilePath , sheet_name="Summary")
        fileHandle.close()
    if LogLevel > 1 :
        fileHandle = pandas.ExcelWriter(fullFilePath,mode='a',engine='openpyxl',if_sheet_exists='new')
        # write Header Informatio
        colHeaderList = [ "SliceType" , "QP" , "BitRate" , "AvgPsnr" , "YPsnr" , "UPsnr" , "VPsnr" , "Gradient" ]
        for encInfo in EncInfoList :
            dataFrame = pandas.DataFrame({ colHeaderList[0] : encInfo.FrameTypeList ,
                                           colHeaderList[1] : encInfo.QPList        ,
                                           colHeaderList[2] : encInfo.BitRateList   ,
                                           colHeaderList[3] : encInfo.YUVPsnrList   ,
                                           colHeaderList[4] : encInfo.YPsnrList     ,
                                           colHeaderList[5] : encInfo.UPsnrList     ,
                                           colHeaderList[6] : encInfo.VPsnrList     ,
                                           colHeaderList[7] : encInfo.GradientList  })
            dataFrame.to_excel(fileHandle , sheet_name=encInfo.SeqName)
        fileHandle.save()
        fileHandle.close()

write_information_filetype = {
    '.log'  : write_information_log  ,
    '.txt'  : write_information_txt  ,
    '.csv'  : write_information_csv  ,
    '.xlsx' : write_information_xlsx ,
}

def write_information( EncInfoList , WriteFilePath , WriteFileType , WriteFileName , LogLevel ) :
    '''
    Write Information from EncInfoList to WriteFileName under WriteFilePath in accroding to WriteFileType
    '''
    create_file( WriteFilePath , WriteFileName , WriteFileType )
    write_information_filetype.get(WriteFileType)( EncInfoList , WriteFilePath , WriteFileName , LogLevel )

def main():
    args = parse_args()
    targetFileList = traverse_files( args.ReadFilePath , args.ReadFileType )
    encInfoList = read_information( targetFileList , args.EncoderName , args.LogLevel )
    encInfoList = delete_nonsequence( encInfoList )
    encInfoList = sort_sequence_ctc.get( args.CtcType )( encInfoList )
    write_information( encInfoList , args.WriteFilePath , args.WriteFileType , args.WriteFileName , args.LogLevel )


if __name__ == '__main__':
    main()


    