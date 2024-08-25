# Créer un "Virtual ENVironment"
Se placer avec la console à l'endroit où l'on veut créer le VENV (à l'aide de la commande "cd"/change directory), dans notre cas à la racine du projet :
```batch
c:\path\to\python.exe -m venv c:\path\to\.venv
```
 
# Activer le VENV (sur Windows)
En étant avec la console à la racine du projet (pour désactiver le VENV : remplacer activate par deactivate):

### Command prompt
```batch
.venv\Scripts\activate
```

### PowerShell
```batch
.venv\Scripts\Activate.ps1
```

# Installer les "requirements"
En étant toujours avec la console à la racine du projet & en ayant activé le VENV :
```batch
pip install -r requirements.txt
```

# Créer un exécutable à partir du code source
En étant toujours avec la console à la racine du projet & en ayant activé le VENV :
```batch
pyinstaller --add-data "Logos;Logos" --add-data "GABARIT.docx;." --add-data "GABARITCOLLABS.docx;." --add-data "GABARIT.xlsm;." --add-data "GABARITCOLLABS.xlsm;." --add-data "INFO.xlsx;." --add-data "maj_gen_py.bat;." --add-data "LICENSE;." --icon=Logos\ETS_Logo.ico --noconsole AutoBibPlus.py
```
Pour les prochains, il suffit d'utiliser le fichier de spécifications qui a été créé (vous pouvez d'ailleurs le modifier) :
```batch
pyinstaller AutoBibPlus.spec
```
Le dossier contenant l'exécutable et toutes les dépendances se trouve alors dans AutoBib\dist\AutoBibPlus !

