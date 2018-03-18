# -*- coding: UTF-8 -*-
'''
Created on 2016年3月9日

@author: pacao
'''
import os
import sys
import re
import shutil
import xlrd
import ConfigParser
from robot.api import TestData

LOGLIST = []
STAMP = True
COUNT = 0
ID_NOT_FOUND = 0

def getInfoFromExcel(excel_path):
    """This KW is used for get instance_id and Tester from excel,return a dict"""
    book = xlrd.open_workbook(excel_path)
    field = book.sheets()[0].row_values(0)
    insid_index = field.index('id')
    tester_index = field.index('Responsible Tester')
    #tester_index = field.index('Tester')
    insid = book.sheets()[0].col_values(insid_index)
    tester = book.sheets()[0].col_values(tester_index)
    insid.pop(0) #delete 'id'
    tester.pop(0) #delete 'tester'
    #print insid
    #print tester
    info = {}
    for i in range(len(insid)):
        info[int(insid[i])] = str(tester[i])
    
    return info

def TraversalScriptPath(apath,dict_info,user_ini):
    """This KW is used for Transform all testcase files in a folder to new QCID of the build.
    This KW will delete all SVN path in parameter 'path'

    | Input Parameters | Man. | Description |
    | apath | Y | testcase files path |
    | excel_path | Y | the excel which contain instance_id and Tester |

    return true or false
    """
    path = apath.replace('\\','/')
#    recordLogsToList(r'Transforming path----%s'%path)
    if not os.path.exists(path):
        recordLogsToList('%s is not exist!' % path)
        return False
    if os.path.isfile(path):
        if '.html' in path:
            parseTestcase(path,dict_info,user_ini)
    elif os.path.isdir(path):
        if '.svn' in path:
            pass
        else:
            searchfile = os.listdir(path)
            for vpath in searchfile:
                childpath = path + '/' + vpath
                TraversalScriptPath(childpath,dict_info,user_ini)
    else:
        recordLogsToList('%s is an unknown object,I can not handle it!' % path)

    return True

def recordLogsToList(log):
    """This KW is used for recordlogs to global log list and print it

    """
    print log
#    global LOGLIST
    LOGLIST.append(log)

def parseTestcase(file_,dict_info,user_ini):
    """This KW is used for change owner tags in a Testcase file"""
    
    reinsid = re.compile('QC_*?([\d-]+)')
    try:
        suite = TestData(source = "%s" %(file_))
    except Exception:
        recordLogsToList('Warning: TestData analyze file [%s] Failed' % file_)
        return False
    for mytestcase in suite.testcase_table:
        if not mytestcase.tags.value:
            recordLogsToList('%s QCID is missed,Please input it in your script!' %file_)
            return False
        for tag in mytestcase.tags.value:
            if 'QC_' in tag:
                insid = reinsid.findall(tag)
                #print insid
                #break
                if dict_info.has_key(int(insid[0])):
                    tester = dict_info[int(insid[0])]
                    print "test:",file_
                    print 'tester:',tester
                    full_name = get_full_name(tester,user_ini)
                    if full_name == 'NONE':
                        break
                    
                    email_name ='Owner-%s@nokia.com' % full_name
                    for ftag in suite.setting_table.force_tags.value:
                        #print ftag
                        if 'Owner-'in ftag or 'owner-' in ftag:
                            owner = re.findall(r'[O|o]wner-(.*)@.*.com',ftag)
                            if owner[0]!=full_name:
                                flag = suite.setting_table.force_tags.value.index(ftag)
                                #print flag
                                suite.setting_table.force_tags.value[flag] = email_name
                                suite.save()
                                recordLogsToList('%s'%file_)
                                recordLogsToList('%s-----> %s' %(ftag,email_name))
                                global STAMP
                                STAMP = False
                                global COUNT
                                COUNT = COUNT+1
                                recordLogsToList('=========%d cases has modified==========' %COUNT)
                else:
                    recordLogsToList('%s'%file_)
                    recordLogsToList('Warning: ID %s not find in excel,please check it!' % int(insid[0]))
                    recordLogsToList('==========================================')
                    global ID_NOT_FOUND
                    ID_NOT_FOUND = ID_NOT_FOUND+1
    
    return True

def get_full_name(short_name,path):
    config = ConfigParser.ConfigParser()
    config.readfp(open(path))
    #print config.sections()
    try:
        
        result = config.get('USER_INFO',short_name)
    except ConfigParser.NoOptionError,e:
        result = 'NONE'
    return result  

def recordLogsToFile(logpath):
    """This KW is used to write global log list to a file
    | Input Parameters | Man. | Description |
    | logpath | Y | The path which is used for save logs |

    """
    ret = True
    global LOGLIST
    if not os.path.exists(logpath):
        os.makedirs(logpath)

    f = open(logpath+'/TesterUpdatelogs.log','wb')
    LOGLIST = [line+'\n' for line in LOGLIST]
    try:
        f.truncate()
        f.writelines(LOGLIST)
    except Exception:
        print 'Write logs to path %s failed!' %logpath
        print Exception
        ret = False
    finally:
        f.close()
    return ret




# if len(sys.argv) == 3:
#     if sys.argv[1] and sys.argv[2]:
#         path = os.path.dirname(sys.argv[0])
#         user_ini = path+'\user.ini'
#         print 'Please wait......'
#         print '=================start===================='
#         dict_info = getInfoFromExcel(sys.argv[2])
#         TraversalScriptPath(sys.argv[1],dict_info,user_ini)
#         print '==================end====================='
#         if ID_NOT_FOUND>0:
#             recordLogsToList('%s instance_id need be checked!'%ID_NOT_FOUND )
#         if STAMP:
#             recordLogsToList('All testers are right,no cases need be modified!')
#         
#         recordLogsToFile(path)
# else:
#     print 'Wrong argv input!'


user_ini = r'D:\workspace\python_test\demo\user.ini'
log_path = r'D:\test'
excel_path = r'D:\userdata\pacao\Desktop\QCInfoTransfer tool\2017-03-09_10_15_21.csv'
dict_info = getInfoFromExcel(excel_path)
file = r'D:\TA_Scripts\TL17SP\DevHZ3\DevHZ3_FV2'
result = TraversalScriptPath(file,dict_info,user_ini)
recordLogsToFile(log_path)
if STAMP:
    recordLogsToList ('No case owner need be modified!')
else:
    recordLogsToList ('Done!')