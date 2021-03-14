# premiere-sequence-to-excel

## Overview

Python script for exporting Adobe Premiere sequence clips into excel file.

![sample image](https://i.loli.net/2021/03/14/afLEAT12ois9uOR.png)

 Testing Environment ：Windows 10 20H2，Adobe Premiere 2020(14.5.0)，Python 3.6



## Usage

### Requirements

```
pymiere==1.2.1
XlsxWriter==0.9.8
```



### Install Pymiere Link extension 

The following method is available for Windows，macOS users can refer to extension installation section [here](https://github.com/qmasingarbe/pymiere/blob/master/README.md#installation).

- Download [pymiere_link.zxp](https://github.com/qmasingarbe/pymiere/blob/master/pymiere_link.zxp) file.

- Unzip it to  `C:\Program Files (x86)\Common Files\Adobe\CEP\extensions`

- Start Premiere，if you can find `Pymiere Link` in `Window → Extensions`, then the installation is successful.

  

### Examples

Start Adobe Premiere and activate a sequence，open a console, change the current working directory to where the `pr_to_excel.py` is located，and type：

```
$ python pr_to_excel.py
```

The script then generates a  `sequence_clip.xlsx` file in the current directory，by default, it will take the preview image from the middle frame of each clip.



You can also customize the output path and whether the preview image is taken from the first frame of the clip(m 0) or from the middle frame(m 1):

```
$ python pr_to_excel.py -o data/test.xlsx -m 1
```

```
$ python pr_to_excel.py -o D:/documents/test.xlsx -m 0
```



  

## Notes

- Timecode sometimes has an error of 1 frame.
- 4k video may make excel file too big.



