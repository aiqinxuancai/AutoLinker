# AutoLinker

e支持库，通过hook CreateFile函数，来实现Linker的实时切换，对应源文件来保存所使用的Linker配置文件，减少设置


## 使用方法
在e目录的AutoLinke\Config中，存放link.ini的各种版本，命名如：vc6.ini vc2017.ini等。
编译后将AutoLinker.fne放在e的lib目录中，并启动候选AutoLinker支持库，你可以在工具条中看到一个新增加的按钮。

