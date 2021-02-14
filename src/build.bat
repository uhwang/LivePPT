del .\LivePPT\liveppt.exe
del .\LivePPT\lib\library.zip

python setup.py build

copy /y newhymal-v.0.0.1.hdb .\LivePPT\newhymal.hdb
copy /y readme.log .\LivePPT
copy /y outline.bas .\LivePPT
copy /y shadow.bas .\LivePPT