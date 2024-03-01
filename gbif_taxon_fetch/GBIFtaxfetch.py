# README
# install  python3 -m pip install requests
#import requests
import urllib.parse,urllib.request, json, sys, time, pickle
import configparser
from pathlib import Path


def readnames(filename): # Read a text file with taxon names, one per line
    print(f"Reading name file {filename}")
    with open(filename) as f:
        d = f.readlines()
        print(f"\tReading name file {filename} done")      
        return d
        
# Convert input files to CSV, store in work_dir
class gbifCache:
    """Simple locally caching data source for GBIF name data. Stores data in a Python binary pickle."""
    def __init__(self,filename):
        self.data = {}
        self.fn = Path(filename)
    def open(self):
        # Read pickled data if exists
        print("Loading GBIF cache data from", self.fn)
        try:
            with open(self.fn,"rb") as f: self.data = pickle.load(f)
        except FileNotFoundError:
            print("No pickle file found")
            self.data = {}
    def flush(self):
        self.data = {}    
    def close(self):        
        with open(self.fn,"wb") as f: pickle.dump(self.data,f)
    def getGBIFData(self,taxonname,classname):# classname = limit name searches to a taxon
        "returns (name,id,family)"
        gbifURLbase = "https://api.gbif.org/v1/species/match?class=%s&name=%s&strict=True"
        gbifURLvalidname = "https://api.gbif.org/v1/species/%i"
        searchname = urllib.parse.quote(taxonname)
        print("Querying GBIF for %s" % searchname)
        if searchname in self.data: # Check if name is in cached GBIF data
            print("Found in cache")
            return self.data[searchname]
        else: # If not in cache, then call GBIF API to check name
            self.data[searchname] = ("","","") # Default value
            time.sleep(1.0) # slow down not to overload the Web query
            GBIFurl = gbifURLbase % (classname,searchname)
            with urllib.request.urlopen(GBIFurl) as f:
                res = json.loads(str(f.read(),"utf-8"))
                if res['matchType'] == 'NONE': return self.data[searchname] # No match
                else: 
                    if (res["synonym"] is True): # lookup recommended name for a synonym
                            validkey = res["speciesKey"]
                            with urllib.request.urlopen(gbifURLvalidname % validkey) as f2:
                                res = json.loads(str(f2.read(),"utf-8"))
                    key = res.get("usageKey","")
                    self.data[searchname] = (res["scientificName"],key,res.get("family",""))
            return self.data[searchname]
        

def ReadConfig(configfile): # Todo, check error states if opening fails (malformatted file & other errors)
    configfile = Path(configfile)
    if not configfile.exists(): return False
    config = configparser.ConfigParser()
    config.read(str(configfile))
    # Check that working directory etc. exist?
    return config

def splitGBIFname(n):
    # Simple implementation, may fail
    # Return name, author
    if not n.strip(): return (n,"") 
    sep = " "
    x = n.split(sep)
    if len(x) > 2:
        name = sep.join(x[0:2])
        author = sep.join(x[2:])
        return (name, author)
    else: return (n,"")

if __name__ == '__main__':
    print("SYSARGV = ", sys.argv)
    configfile = r"P:\h978\insect\Kahanpää\SMALL_TOOLS\gbif_taxon_fetch\mzhconfig.ini" # See function ReadConfig below for default config
    cachename = "P:\h978\insect\Kahanpää\SMALL_TOOLS\gbif_taxon_fetch\gbifcache.dat"
    if  (len(sys.argv) > 1):
        inputfn = sys.argv[1]
    else:
        print("No input data file, exiting...")
        sys.exit()
    outputfn = inputfn + ".authors.txt"
    taxnames = readnames(inputfn)
    # Create list of unique names
    taxnames = {x:"" for x in taxnames}.keys()
    print("Reading config file")
    conf = ReadConfig(configfile)
    if not conf:
        print("Config file %s not found" % configfile)
        sys.exit()
    if conf["DEFAULT"].getboolean("flush_GBIF_cache"):        
        print("Flushing GBIF cache ... ",end='', flush=True)
        try:
            gbifsource = gbifCache("gbifcache.dat")
            gbifsource.open()
            gbifsource.flush()
            print("clear")
        finally:
            gbifsource.close()            
    gbifsource = gbifCache(cachename)
    gbifsource.open()
    with open(outputfn,"w") as of:
        # Header line
        of.write("original name; GBIF name; GBIF author; GBIF family\n")
        try:
            for n in taxnames:
                n = n.strip()
    #            print("Checking name %s" % n)
                result = gbifsource.getGBIFData(n,conf["DEFAULT"]["class"])
                # Result = (fullname, GBIF_id, family)
                if result[0]: #Has some output
                    of.write("%s;%s;%s;%s\n" % ( \
                        n,
                        splitGBIFname(result[0])[0], \
                        splitGBIFname(result[0])[1], \
                        result[2]
                        ) )

        except Exception as err: # Catch any Exception
            if 'gbifsource' in dir(): gbifsource.close()
            raise err
    gbifsource.close()

    # 1. GET TAXON NAMES FROM EXCEL SHEET

    # 2. CALL LAJITIETOKESKUS API (need permission)

    # 3. CALL GBIF API for laji.fi failed cases 

