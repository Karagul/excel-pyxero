# excel-pyxero
Accessing Xero through Excel using xlwings and pyxero

### You will need/dependencies

[Anaconda](https://www.continuum.io/downloads) - This has only been tested with Python 3

[xlwings](https://github.com/ZoomerAnalytics/xlwings)

[pyxero](https://github.com/freakboy3742/pyxero)

### Initial Setup
Once anaconda is installed run the following commands 
```
python –m pip install xlwings
python –m pip install xlwings --upgrade

python –m pip install comtypes
python –m pip install comtypes --upgrade

python –m pip install pyxero
python –m pip install pyxero --upgrade
```

You'll also need open Excel and:

File > Options > Trust Centre > Trust Centre Settings > Macro Settings > Enable “Trust access to the VBA project object model”

Additionally you will need to register a private application with Xero for your organisation. [This](https://developer.xero.com/documentation/auth-and-limits/private-applications) details the steps required to generate the certificate and register it for use with Xero.
