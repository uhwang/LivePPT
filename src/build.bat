del .\LivePPT\liveppt.exe
del .\LivePPT\lib\library.zip

python setup.py build

copy /y newhymal-v.0.0.1.hdb .\LivePPT\newhymal.hdb
copy /y readme.log   .\LivePPT
copy /y default.pptx .\LivePPT
copy /y shadow.bas   .\LivePPT
copy /y shadow.bas   .\LivePPT