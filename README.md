# premiere-sequence-to-excel

## 介绍

一个将 Pr 当前序列上的片段信息导出为 excel 文件的 Python 脚本。

![sample image](https://i.loli.net/2021/03/14/afLEAT12ois9uOR.png)

测试环境：Windows 10 20H2，Adobe Premiere 2020(14.5.0)，Python 3.6





## 使用

### 依赖

```
pymiere==1.2.1
XlsxWriter==0.9.8
```



### 安装 Pymiere Link 插件

以下为 Windows 系统中的操作方法，macOS 用户可参见[这里](https://github.com/qmasingarbe/pymiere/blob/master/README.md#installation)的插件安装部分。

- 下载 [pymiere_link.zxp](https://github.com/qmasingarbe/pymiere/blob/master/pymiere_link.zxp) 文件

- 使用解压缩软件将其解压到 `C:\Program Files (x86)\Common Files\Adobe\CEP\extensions`

- 打开 Pr，如果可以在`窗口 → 扩展`中找到`Pymiere Link`，便说明安装成功

  

### 示例

启动 Adobe Premiere，并激活一个序列，在 cmd 中转到 `pr_to_excel.py` 所在文件夹，直接输入：

```
$ python pr_to_excel.py
```

脚本便会在当前目录下生成一个 `sequence_clip.xlsx` 文件，默认从每个片段中间帧截取预览图。



也可以自定义输出路径以及从片段的第一帧（m 为 0）还是中间帧（m 为 1）截取预览图：

```
$ python pr_to_excel.py -o data/test.xlsx -m 1
```

```
$ python pr_to_excel.py -o D:/documents/test.xlsx -m 0
```



  

## 备注

- 目前还没有做不同情况下的测试
- 时间码有时会出现 1 帧的误差
- 4k 视频可能会使 excel 文件过大



