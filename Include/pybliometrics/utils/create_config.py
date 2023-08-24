import configparser, os
from typing import List, Optional
from pathlib import Path

def create_config(keys: Optional[List[str]] = None,
                  insttoken: Optional[str] = None,
                  docs_path: str = Path(os.path.expanduser("~")) / "Documents"):
    """Initiates process to generate configuration file.

    :param keys: If you provide a list of keys, pybliometrics will skip the
                 prompt.  It will also not ask for InstToken.  This is
                 intended for workflows using CI, not for general use.
    :param insttoken: An InstToken to be used alongside the key(s).  Will only
                      be used if `keys` is not empty.
    """
    from .constants import CONFIG_FILE, DEFAULT_PATHS

    config = configparser.ConfigParser()
    config.optionxform = str

    # Set directories
    config.add_section('Directories')
    for api, path in DEFAULT_PATHS.items():
        config.set('Directories', api, str(path))

    # Set authentication for Scopus and SciVal
    section_name='Authentication'
    config.add_section(section_name)
    if keys:
        if not isinstance(keys, list):
            raise ValueError("Parameter `keys` must be a list.")
        key = ", ".join(keys)
        token = insttoken

    config.set(section_name, 'APIKey', key)

    if token:
        config.set(section_name, 'InstToken', token)

    # Set default values
    config.add_section('Requests')
    config.set('Requests', 'Timeout', '20')
    config.set('Requests', 'Retries', '5')

    # Définir le chemin dans le fichier de configuration
    config['Docs Path'] = {
        'Path': str(docs_path)
    }

    # Write out
    CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_FILE, "w") as ouf:
        config.write(ouf)
    # print(f"Configuration file successfully created at {CONFIG_FILE}\n"
    #       "For details see https://pybliometrics.rtfd.io/en/stable/configuration.html.")
    return config
