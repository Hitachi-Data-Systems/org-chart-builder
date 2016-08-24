# org-chart-builder
Tool to generate Org Chart PPTs from a source spreadsheet


Get the org-chart-builder code from github:
https://github.com/Hitachi-Data-Systems/org-chart-builder

I ran "sudo python setup.py install" on the setup.py file in the org-chart-builder directory, but I'm not sure if that was strictly necessary.

Install pip (taken from http://stackoverflow.com/questions/1231688/how-do-i-remove-packages-installed-with-pythons-easy-install)
$ wget https://bootstrap.pypa.io/get-pip.py
$ python get-pip.py

NOTE: If you don't have wget, you can just wget the file on another machine and copy it over

sudo pip install jdcal
sudo pip install lxml
sudo pip install python-pptx

That took care of package installation.

To run the program, from the org-chart-builder directory, run something like:
python ppt_draw.py -h

python ppt_draw.py /path/to/spreadsheet.xlsm -o /path/to/output/OrgChart.pptx -t
