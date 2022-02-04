### Instruction
1. Create/choose your python environment.
2. Install python requirements:  
```shell
pip install -r requirements.txt
```
3. Clean your Excel template from the old data and save it as '.xlsx' in ./template
4. Put your csv files in ./input_csv
5. Create ./output_excel dir for output files.
6. Set correct directiries, template name and other params in **run.py**
7. Run the programm:
```shell
python run.py
```
The processing of one files might take about 30-40 seconds.

## Known issues
Some files can't be processed due to their special aspects.
For 212 csv files provided, 6 was not processed and should be populated manually.
The issue will be fixed as soon as possible.