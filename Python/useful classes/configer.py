import os
import configparser

class Config(configparser.ConfigParser):
    """
    Creates a configfile <filepath> with the structure
    contained in <default_dict> as default values.
    
    <default_dict> should look like this:
    default_dict = {
        "GENERAL": {
            "theme": 1,
            "splash_screen": True,
            ...
        },
        "NETWORK": {
            "server_addr": "127.0.0.1",
            "server_port": "1337",
            ...
        },
        ...
    }
    
    The configfile will be created upon instantiation.
    The method <get_config()> can be used to retrieve
    values from the config in real time from the file.
    This means, if the config file is changed, e.g. by
    the user, during runtime the program will fetch
    these new values.
    """
    def __init__(self, filepath, default_dict={}):
        super().__init__()
        self.filepath = filepath
        if not os.path.isfile(filepath):
            try:
                for section in default_dict.keys():
                    self.add_section(section)
                    for item, value in default_dict[section].items():
                        self[section][item] = str(value)
                with open(filepath, "w") as configfile:
                    self.write(configfile)
            except Exception as e:
                raise Exception(e.__class__, "Parameter default_dict likely did not have the required structure.")
    
    def get_option(self, itemname, cast=str, fallback=None):
        assert isinstance(itemname, str)
        assert cast in (str, int, float, bool)
        self.clear()
        self.read(self.filepath)
        found = False
        for s in self.sections():
            for c in self.options(s):
                #print(c, itemname, c == itemname)
                if c == itemname:
                    value = self[s][c]
                    #print(value)
                    found = True
                    break
                
        # If setting does not exist, return fallback value
        if not found:
            return fallback
        
        # Cast the retrieved value if demanded
        if type(value) != cast:
            try:
                return cast(value)
            except Exception as e:
                return fallback
        else:
            return value
