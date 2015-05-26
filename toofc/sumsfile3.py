import multiprocessing
from risk.zinc.api.zn import ZnApi

def execute_query1():
    print 'inside execute query'
    zn_nocache = ZnApi('universal', 'qa', use_cache=False)
    date1 ='2015-05-19'
    print date1
    snapshot_id = zn_nocache.snapshot_id(date1, 'EOD')
    select_fields = ['SOURCE','SUM("MTM")']
    qzt_qa1 = zn_nocache.query(snapshot_id, select_fields)
    return qzt_qa1
    
def execute_query2():
    pass
    
def execute_query3():
    pass
    
def runner(arg, qa='homedirs/home/ZincQA/python;ps', prod='ps'):
    out_file_name, query = arg
    query1 = execute_query1()
    # Create an output file in the desired place
    out_file = open(out_file_name, "w")
    out_file.write(query1)
    out_file.close()
    # return out_file_name


def main():
    
    # list tuples of desired name of output filename, query
    outlist = ["out1", "out2", "out3"]  
    queryfunclist = [execute_query1, execute_query2, execute_query3]
    args = zip(outlist, queryfunclist)
    
    for i in args:
        # Create a pool of subprocesses
        pool = multiprocessing.Pool(processes=2)
        pool.map(runner, args,1)
