# -*- coding=utf-8 -*-
# author: yanyang.xie@gmail.com

from xml.etree.ElementTree import ElementTree, Element
import xml.etree.ElementTree as ET

def read_xml(xml_file_path):
    '''
    Read xml file from local file
    @param xml_file_path: local xml file path
    @return: ElementTree
    '''
    tree = ElementTree()
    root = tree.parse(xml_file_path)
    return tree

def read_xml_from_string(xml_string):
    '''
    Read xml file from string
    @param xml_string: xml format string
    @return: root Element
    '''
    root = ET.XML()
    return root

def find_sub_element(parent_element, element_tag_name):
    '''
    Find the first node in the parent element
    @param parent_element: Element
    @param element_tag_name: Element name
    '''
    return parent_element.find(element_tag_name)

def find_sub_elements(parent_element, element_tag_name):
    '''
    Find all the node in the parent element
    @param parent_element: Element
    @param element_tag_name: Element name
    '''
    return parent_element.findall(element_tag_name)

def find_elements(element_tree, path):
    '''
    Find all the node in the tree using xpath
    @param element_tree: ElementTree
    @param path: xpath
    '''
    return element_tree.findall(path)

def find_elements_by_attributes(element_list, expected_key_value_dict):
    '''
    Find nodes which has special attribute value
    @param element_list:
    @param expected_key_value_dict:expected attribute key/value dict
    ''' 
    result_elements = []
    for element in element_list:
        if has_element_attribute(element, expected_key_value_dict):
            result_elements.append(element)
    return result_elements

def has_element_attribute(xml_element, expected_key_value_dict):
    '''
    Check whether element has special attribute or not
    @param xml_element: Element
    @param expected_key_value_dict: attribute key/value dict
    '''
    for key in expected_key_value_dict:
        if xml_element.get(key) != expected_key_value_dict.get(key): 
            return False
    return True
 
def change_element_properties(element_list, dest_key_value_dict, is_delete=False):
    '''
    Change or add or delete element attribute
    @param element_list:
    @param dest_key_value_dict: changed attribute key/value dict
    @param is_delete: if is_delete is True, will delete its attribute
    ''' 
    for element in element_list:
        for key in dest_key_value_dict:
            if is_delete:
                if key in element.attrib:
                    del element.attrib[key]
            else:
                element.set(key, dest_key_value_dict.get(key))

def change_element_text(element_list, text, is_append=False, is_delete=False):
    '''
    Change or add or delete element text
    @param element_list:
    @param text: changed element text
    @param is_append: if is_append is True, will append its text
    @param is_delete: if is_delete is True, will delete its text
    '''
    for element in element_list:
        if is_append:
            element.text += text
        elif is_delete:
            element.text = ""
        else:
            element.text = text

def create_new_element(element_tag, attribute_dict={}, element_text=''):
    '''
    Create a new element object
    @param element_tag: element name
    @param attribute_dict: element attibutes
    @param element_text: element text
    '''
    element = Element(element_tag, attribute_dict)
    element.text = element_text
    return element

def add_sub_element(par_element_list, sub_element):
    for par_element in par_element_list:
        par_element.append(sub_element)

def delete_sub_element(element_list, sub_element_tag, kv_map={}):
    for par_element in element_list:
        children = par_element.getchildren()
        for child in children:
            if child.tag == sub_element_tag:
                if len(kv_map) > 0:
                    if has_element_attribute(child, kv_map):
                        par_element.remove(child)
                else:
                    par_element.remove(child)
                    
def write_xml(element_tree, file_path, encoding='utf-8', xml_declaration=True, default_namespace=None):
    '''
    Write element tree into local file
    @param element_tree: ElementTree
    @param out_path: local file
    '''
    element_tree.write(file_path, encoding=encoding, xml_declaration=xml_declaration, default_namespace=default_namespace)
