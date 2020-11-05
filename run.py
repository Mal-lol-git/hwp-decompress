import os
import sys
import zlib
import pyzipper 
import olefile
from settings import *

hwp = []
hex_list = []
bin_list = {}

FOL_PATH = sys.argv[1]


def doc_scan(fol_path):
    doc_file = os.listdir(fol_path)
    if len(doc_file) != 0:
        for filename in doc_file:
            if filename != 'result.txt':
                hwp.append(filename)
    else:
        print('please put file on %s folder ' % fol_path)
        sys.exit(1)
        
def ole_stream_data(filename):
    try:
        file_path = os.path.join(FOL_PATH, filename)
        with olefile.OleFileIO(file_path) as ole:
            for row in ole.listdir():
                if row[0] == 'BinData':
                    data = ole.openstream(row[0]+'/'+row[1]).read()
                    bin_list[row[1]] = data          
    except Exception as e:   
        print(e)

def bin_scan(rels_path):
    bin_file = os.listdir(rels_path)
    for filename in bin_file:
        bin_path.append(filename)

def hwp_decompress(key, obj_value, hwp_name):
    try:
        z_obj = zlib.decompressobj(-zlib.MAX_WBITS)
        result = z_obj.decompress(obj_value)
  
        folder_path = os.path.join(SAVE_PATH)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
    
        save_file_path = os.path.join(folder_path, hwp_name + '_' + key + '.decompress')  
        f = open(save_file_path, 'wb')
        f.write(result)
        f.close()
        
    except Exception as e: 
        result = obj_value.decode("ascii", errors="ignore")
        
        folder_path = os.path.join(SAVE_PATH)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
    
        save_file_path = os.path.join(folder_path, hwp_name + '_' + key + '.hex')  
        f = open(save_file_path, 'w')
        f.write(result)
        f.close()   
   
def malscan():
    for name in hwp:
        ole_stream_data(name)

        for key in bin_list.keys():
            value = bin_list.get(key)
            hwp_decompress(key, value, name)
        bin_list.clear()
   
doc_scan(FOL_PATH)
malscan()



