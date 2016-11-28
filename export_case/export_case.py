'''
Created on Aug 26, 2016

@author: Vicky
'''

import os
import sys
import re

from excel_util import ExcelUtil
from xml_util import read_xml
import string
import xml_util
import xlwt
import getopt

head_row = 0
case_num_col = 0
sum_col = 1
des_col = 2
steps_col = 3
sub_steps_col = 4
request_col = 5
check_point_col = 6
case_name_col = 7

supported_samplers = ['HTTPSamplerProxy','SystemSampler','BeanShellSampler','IfController','WhileController','org.apache.jmeter.protocol.mongodb.sampler.script.Sampler','MongoScriptSampler','BeanShellPreProcessor','BeanShellPostProcessor']
controllers_list = ['IfController','WhileController']
samplers_list = ['HTTPSamplerProxy','SystemSampler','BeanShellSampler','MongoScriptSampler','BeanShellPreProcessor']
supported_groups = ['ThreadGroup','SetupThreadGroup' ,'PostThreadGroup']
supported_assertion = ['ResponseAssertion','XPathAssertion','BeanShellAssertion','com.atlantbh.jmeter.plugins.jsonutils.jsonpathassertion.JSONPathAssertion']
script_dir = '' 
case_dir = ''
        
header_style = xlwt.easyxf('font: color-index white, height 280, name Arial, bold True; align: vertical center, horizontal center; pattern: pattern solid, fore_colour sea_green; borders: left 1, right 1, bottom 1, left_colour white, right_colour white, bottom_colour white;')
cell_style = xlwt.easyxf('font: color-index black, height 240, name Arial; align: wrap on, horizontal left, vertical top;')
des_style = xlwt.easyxf('font: color-index black, height 240, name Arial; align: wrap on, horizontal center, vertical center;')
num_style = xlwt.easyxf('font: color-index black, height 240, name Arial; align: wrap on, horizontal center, vertical top; ')

class Case:
    def __init__(self):
        self.case_summary = ''
        self.case_description = ''
        self.steps = [] # Class Step
        self.case_name = ''
        
    def add_step(self, step):
        self.steps.append(step)

class SubStep:
    def __init__(self):
        self.request = ''
        self.sub_step = ''
        self.check_points = []
    
class Step:
    def __init__(self):
        self.sub_steps = []  # Class SubStep
        self.step_des = ''
        
    def add_sub_step(self, sub_step):
        self.sub_steps.append(sub_step)
    
class KeyHashTree:
    def __init__(self):
        self.key = []  # it could be thread group or all kinds of samplers
        self.hash_tree = []
 
def convert_to_hash_tree(xml_list,supported_tags):
    key_hash_tree = KeyHashTree()
    for i in range(0,len(xml_list)):    
        if(xml_list[i].tag in supported_tags):
            key_hash_tree.key.append(xml_list[i])
            key_hash_tree.hash_tree.append(xml_list[i+1]) 
    return key_hash_tree  
                  
def get_sub_dirs(script_dir):
    
    if script_dir=='':
        script_dir=os.getcwd()
        
    return os.listdir(script_dir)


def get_all_scripts(sub_dir):
         
    all_files=[]
    file_names=os.listdir(sub_dir)
    
    for file_name in file_names:
        if file_name.endswith('.jmx'):
            full_file_name = os.path.join(sub_dir, file_name)
            all_files.append(full_file_name)
                
    return all_files

def is_enabled (xml_element):
    if xml_element.get('enabled') =='true':
        return True
    return False

def combine_step (number_str,xml_element):
    return number_str+'. '+xml_element.get('testname') 
    

def combine_request(controller_request):
    if controller_request.tag == 'HTTPSamplerProxy':
        return send_request_link(xml_util.find_sub_elements(controller_request,'stringProp'))     
    elif  controller_request.tag == 'org.apache.jmeter.protocol.mongodb.sampler.script.Sampler' or controller_request.tag == 'MongoScriptSampler' :
        return send_mongo_info(xml_util.find_sub_elements(controller_request,'stringProp'))
    elif  controller_request.tag == 'SystemSampler' :
        return send_os_command(xml_util.find_sub_elements(controller_request,'stringProp'))  
    elif  controller_request.tag == 'BeanShellPreProcessor' :
        return send_beanshell_processor_info(xml_util.find_sub_elements(controller_request,'stringProp'))
    elif  controller_request.tag == 'BeanShellSampler' :
        return send_beanshell_sampler_info(xml_util.find_sub_elements(controller_request,'stringProp'))  
    else :
        return ''  


def add_sub_steps (ex_step,sub_steps,check_point_list):
    sub_number = 1
    for x in range(0, len(sub_steps)) :
                       
        if is_enabled(sub_steps[x]):           
            if sub_steps[x].tag in controllers_list:
                controller_ex_sub_step = SubStep() 
                controller_ex_sub_step.sub_step = combine_step(str(sub_number),sub_steps[x])
                ex_step.add_sub_step(controller_ex_sub_step) 
                
                controller_requests_tree = xml_util.find_elements(check_point_list[x],'*')  
                controller_requests = convert_to_hash_tree(controller_requests_tree,supported_samplers).key
                controller_check_point_list = convert_to_hash_tree(controller_requests_tree,supported_samplers).hash_tree
                controller_sub_number = 1
                for i in range(0,len(controller_requests)):                    
                    if is_enabled(controller_requests[i]):
                        ex_sub_step = SubStep()
                        ex_sub_step.sub_step = combine_step(str(sub_number)+'.'+str(controller_sub_number),controller_requests[i]) 
                        controller_sub_number +=1
                        
                        ex_sub_step.request = combine_request(controller_requests[i])                        
                                                                 
                        check_point_num = 1
                        for assertion_number in range(0, len(supported_assertion)):
                            check_points =  xml_util.find_sub_elements(controller_check_point_list[i], supported_assertion[assertion_number])  
                            for m in range(0, len(check_points)):
                                if is_enabled(check_points[m]): 
                                    ex_sub_step.check_points.append(combine_step(str(check_point_num), check_points[m]))
                                    check_point_num +=1
                        ex_step.add_sub_step(ex_sub_step)   
            else: 
                ex_sub_step = SubStep()                 
                ex_sub_step.sub_step = combine_step(str(sub_number),sub_steps[x]) 
                ex_sub_step.request = combine_request(sub_steps[x])
                check_point_num = 1 
                for assertion_number in range(0, len(supported_assertion)):
                    check_points =  xml_util.find_sub_elements(check_point_list[x], supported_assertion[assertion_number])
                    for m in range(0, len(check_points)):
                        if is_enabled(check_points[m]): 
                            ex_sub_step.check_points.append(combine_step(str(check_point_num), check_points[m]))
                            check_point_num +=1
                        
                ex_step.add_sub_step(ex_sub_step)   
            sub_number += 1  
          
def send_request_link(request_parameters):
    request_domain = ''
    request_port = ''
    request_paras = ''
    request_type = ''
    for i in range(0, len(request_parameters)):
        if request_parameters[i].get('name') == 'HTTPSampler.domain' and request_parameters[i].text is not None:
            request_domain = request_parameters[i].text
        elif request_parameters[i].get('name') == 'HTTPSampler.port' and request_parameters[i].text is not None:
            request_port = request_parameters[i].text
        elif request_parameters[i].get('name') == 'HTTPSampler.path' and request_parameters[i].text is not None:
            request_paras = request_parameters[i].text
        elif request_parameters[i].get('name') == 'HTTPSampler.method' and request_parameters[i].text is not None:
            request_type = request_parameters[i].text  
                    
    return request_type +': http://' + request_domain + ":" + request_port + request_paras
    
def send_os_command(request_parameters):
    os_command = ''
    for i in range(0, len(request_parameters)):
        if request_parameters[i].get('name') == 'SystemSampler.command' and request_parameters[i].text is not None:
            os_command = request_parameters[i].text                       
    return 'OS Command :' + os_command
    
def send_mongo_info(request_parameters):
    database = ''
    password = ''
    source = ''
    username = ''
    script = ''
    
    for i in range(0, len(request_parameters)):
        if request_parameters[i].get('name') == 'database' and request_parameters[i].text is not None:
            database = request_parameters[i].text
        elif request_parameters[i].get('name') == 'password' and request_parameters[i].text is not None:
            password = request_parameters[i].text
        elif request_parameters[i].get('name') == 'source' and request_parameters[i].text is not None:
            source = request_parameters[i].text
        elif request_parameters[i].get('name') == 'username' and request_parameters[i].text is not None:
            username = request_parameters[i].text
        elif request_parameters[i].get('name') == 'script' and request_parameters[i].text is not None:
            script = request_parameters[i].text                       
    return 'Database: ' + database+ ',' + ' Source: ' + source + ','+ ' UserName: ' + username +',' + ' Password: ' + password + ','+ ' Script: ' + script

def send_beanshell_sampler_info(request_parameters):
    bean_shell_query = ''
    bean_shell_file = ''
    bean_shell_parameters = ''
    reset_interpreter = ''
    
    for i in range(0, len(request_parameters)):
        if request_parameters[i].get('name') == 'BeanShellSampler.query' and request_parameters[i].text is not None:
            bean_shell_query = request_parameters[i].text
        elif request_parameters[i].get('name') == 'BeanShellSampler.filename' and request_parameters[i].text is not None:
            bean_shell_file = request_parameters[i].text
        elif request_parameters[i].get('name') == 'BeanShellSampler.parameters' and request_parameters[i].text is not None:
            bean_shell_parameters = request_parameters[i].text
        elif request_parameters[i].get('name') == 'BeanShellSampler.resetInterpreter' and request_parameters[i].text is not None:
            reset_interpreter = request_parameters[i].text
                     
    return 'Reset Interpreter: ' + reset_interpreter+ ',' + ' BeanShell Parameters: ' + bean_shell_parameters + ','+ ' BeanShell file: ' + bean_shell_file +',' + ' BeanShell Query: ' + bean_shell_query
  
def send_beanshell_processor_info(request_parameters):
    bean_shell_script = ''
    bean_shell_file = ''
    bean_shell_parameters = ''
    reset_interpreter = ''
    
    for i in range(0, len(request_parameters)):
        if request_parameters[i].get('name') == 'script' and request_parameters[i].text is not None:
            bean_shell_script = request_parameters[i].text
        elif request_parameters[i].get('name') == 'filename' and request_parameters[i].text is not None:
            bean_shell_file = request_parameters[i].text
        elif request_parameters[i].get('name') == 'parameters' and request_parameters[i].text is not None:
            bean_shell_parameters = request_parameters[i].text
        elif request_parameters[i].get('name') == 'resetInterpreter' and request_parameters[i].text is not None:
            reset_interpreter = request_parameters[i].text
                     
    return 'Reset Interpreter: ' + reset_interpreter+ ',' + ' BeanShell Parameters: ' + bean_shell_parameters + ','+ ' BeanShell file: ' + bean_shell_file +',' + ' BeanShell Script: ' + bean_shell_script
     
def add_steps (case, steps_list,sub_steps_list):
    step_number = 1   
    for i in range(0, len(steps_list)) :
        ex_step = Step()
        if is_enabled(steps_list[i]):
            step = combine_step(str(step_number),steps_list[i])
            
            step_number +=1            
            sub_hash_tree = xml_util.find_elements(sub_steps_list[i],'*')               
            
            sub_steps = convert_to_hash_tree(sub_hash_tree, supported_samplers).key
            check_point_list = convert_to_hash_tree(sub_hash_tree, supported_samplers).hash_tree
            
            add_sub_steps(ex_step,sub_steps,check_point_list)                  
            ex_step.step_des = step
            case.add_step(ex_step)        
    return case
  
def parse_script(file_name): 
    
    case = Case() 
    case.case_name = file_name
    
    try:
        tree = read_xml(file_name)
    except Exception, e: 
        print "Error:cannot parse file: " + file_name
        sys.exit(e)     
    
    testplan = xml_util.find_elements(tree,'.//TestPlan')[0]
    case.case_summary = testplan.get('testname') #case summary
    case.case_description = xml_util.find_sub_element(testplan, 'stringProp').text
    
    all_hash_tree = xml_util.find_elements(tree,'./hashTree/hashTree')[0]
    xml_list = xml_util.find_elements(all_hash_tree,'*')   
    
    steps_list = convert_to_hash_tree(xml_list,supported_groups).key
    sub_steps_list = convert_to_hash_tree(xml_list,supported_groups).hash_tree
    
    add_steps(case, steps_list, sub_steps_list)              
    return case

def output_excel(parent_dir, case_dir):
    
    excel_util = ExcelUtil()
    
    sub_dirs = get_sub_dirs(script_dir)
    
    for sub_dir in sub_dirs:
        if sub_dir.find('.')>-1:
            continue
        case_list = []
        script_list = get_all_scripts(os.path.join(script_dir,sub_dir))
        
        for script in script_list:
            case_list.append(parse_script(script))
        
        case_file = string.split(parent_dir,'/')[-1]+'.xls'   
    
        work_sheet = excel_util.add_sheet(sub_dir)    
        # case header
        work_sheet.row(head_row).set_style(header_style)  
     
        work_sheet.col(case_num_col).width = 3000
        work_sheet.col(sum_col).width = 6000
        work_sheet.col(des_col).width = 6000
        work_sheet.col(steps_col).width = 8000
        work_sheet.col(sub_steps_col).width = 16000
        work_sheet.col(request_col).width = 16000
        work_sheet.col(check_point_col).width = 16000
        work_sheet.col(case_name_col).width = 8000  
         
        work_sheet.row(head_row).set_style(xlwt.easyxf('font: height 500;'))
               
        work_sheet.write(head_row, case_num_col, u'Case #', header_style)
        work_sheet.write(head_row, sum_col, u'Case Summary', header_style)
        work_sheet.write(head_row, des_col, u'Case description', header_style)
        work_sheet.write(head_row, steps_col, u'Steps', header_style)
        work_sheet.write(head_row, sub_steps_col, u'Sub steps', header_style)
        work_sheet.write(head_row, request_col, u'Request', header_style)
        work_sheet.write(head_row, check_point_col, u'Expect Results', header_style)
        work_sheet.write(head_row, case_name_col, u'Case in Jmeter', header_style)    
        
        work_sheet.set_horz_split_pos(head_row+1)
#       work_sheet.set_vert_split_pos(1)
        work_sheet.panes_frozen = True
        work_sheet.remove_splits = True
        
        case_start_row = 1
        case_num = 1
        
        for i in range(0,len(case_list)):
            
            indi_start_row = case_start_row
            excl_case = case_list[i] 
           
            case_sum = excl_case.case_summary
            case_des = excl_case.case_description
            case_name = string.split(excl_case.case_name,'/')[-1]
            case_steps = excl_case.steps

            for j in range(0, len(case_steps)) :
                step_start_row = case_start_row
                step_des = case_steps[j].step_des
                case_sub_steps = case_steps[j].sub_steps
                
                for m in range(0, len(case_sub_steps)):
                    sub_step_start_row = case_start_row
                    sub_step = case_sub_steps[m].sub_step
                    check_point_list = case_sub_steps[m].check_points
                    sub_request = case_sub_steps[m].request
                    
                    if len(check_point_list) == 0:
                        case_start_row += 1
                    else:
                        for n in range(0, len(check_point_list)):
                            check_point = check_point_list[n]
                            work_sheet.write(case_start_row, check_point_col, check_point, cell_style)
                            case_start_row += 1
                    
                    if sub_step_start_row == case_start_row-1:
                        work_sheet.write(sub_step_start_row, sub_steps_col, sub_step, cell_style)
                        work_sheet.write(sub_step_start_row, request_col, sub_request, cell_style)                        
                    elif sub_step_start_row <= case_start_row-1:
                        work_sheet.write_merge(sub_step_start_row, case_start_row-1, sub_steps_col, sub_steps_col, sub_step, cell_style)
                        work_sheet.write_merge(sub_step_start_row, case_start_row-1, request_col, request_col, sub_request, cell_style)                                            
 
                work_sheet.write_merge(step_start_row, case_start_row-1, steps_col, steps_col, step_des, cell_style)
            
            work_sheet.write_merge(indi_start_row, case_start_row-1, case_num_col, case_num_col, case_num, num_style)  
            work_sheet.write_merge(indi_start_row, case_start_row-1, sum_col, sum_col, case_sum, cell_style)  
            work_sheet.write_merge(indi_start_row, case_start_row-1, des_col, des_col, case_des, cell_style)  
            work_sheet.write_merge(indi_start_row, case_start_row-1, case_name_col, case_name_col, case_name, cell_style)  
            case_num +=1       
    case_dir = case_dir if case_dir[-1] == os.sep else case_dir + os.sep        
    excel_util.save(case_dir+case_file)
    print 'done :'+case_dir+case_file

def export_case(script_dir, case_dir):
    
    sub_dirs = get_sub_dirs(script_dir)
    for sub_dir in sub_dirs:
        if sub_dir.find('.')>-1:
            continue
        case_list = []
        script_list = get_all_scripts(os.path.join(script_dir,sub_dir))
        
        for script in script_list:
            case_list.append(parse_script(script))
        
        output_excel(script_dir, sub_dir, case_dir, case_list)
 
def read_parameters():
    global script_dir, case_dir    
    opt_dict = read_opts(['-h', '-f', '-t'])[0]
    if opt_dict.has_key('-h'):
        usage()
    else:
        if opt_dict.has_key('-f'):
            script_dir = string.strip(opt_dict['-f'])
        if opt_dict.has_key('-t'):
            case_dir = string.strip(opt_dict['-t'])
        return True

def check_parameters():
    if script_dir is None or script_dir == '':
        print 'Where is jmeter scripts? Come on, tell me the directory! \n-h: help message'
        return False
    if case_dir is None or case_dir == '':
        print 'Where should the tool put the test case in? Come on, tell me the directory! \n-h: help message'
        return False
    return True

def usage():
#   print '*' * 100
    print 'Usage:'
    print '-h: help message.'
    print '-f: Jmeter script directory.'
    print '-t, Excel case output directory.'
    print 'Example: python exportCase.py -f /root/jmeter/scripts -t /root/case'
#   print '*' * 100
    
def read_opts(short_param_list=[], long_param_list=[]):
    short_params = 'h'
    for param in short_param_list:
        short_params += param.replace('-', '').strip() + ':'
    
    long_params = ["help"]
    for param in long_param_list:
        long_params.append(param.replace('-', '').strip() + '=')
        
    opts, args = getopt.getopt(sys.argv[1:], short_params, long_params)
    
    opt_dict = {}
    for opt, value in opts:
        opt_dict[opt] = value
    
    return opt_dict, args
    
def refine_dir(dir):  
    if len(re.findall(r'(?=.)',dir)) == 1 or len(re.findall(r'(?=./)',dir)) == 1:
        dir = dir.replace('.',os.path.abspath('.'))
    if len(re.findall(r'(?:..|../)',dir)) > 0: #or len(re.findall(r'(?=../)',dir)) > 0:
        dir = dir.replace('..',os.path.abspath('..'))
    dir = dir if dir[-1] != os.sep else dir[0:-1]     
    return dir
         
def main():
    if not read_parameters():
        sys.exit(0)
    
    if not check_parameters():
        usage()
        sys.exit(0)         
    output_excel(refine_dir(script_dir), refine_dir(case_dir))

    
if __name__ == '__main__':
    main()
    
    
        
   
