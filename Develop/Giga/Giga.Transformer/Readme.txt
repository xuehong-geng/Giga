！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！
！Giga.Transformer库是用来处理不同格式数据的读取的工具库。
！
！版本号：1.2
！作者：耿学鸿
！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！

功能：
* 通过配置文件来配置数据读取的模版。
* 使用泛型直接将文件中的数据读到实体对象中。
* 目前支持Excel数据的读取，可读以下模式的数据：
---- 表格数据。一条数据占用1到多行。每条数据的高度必须一致且数据之间没有空行。
---- 表单数据。一张worksheet中包含一个表单对象的数据。
---- 嵌套数据。一张表单中嵌套子表格数据。

使用方法：
1. 将Giga.Transformer.dll添加到工程的应用中。
2. 在项目的配置文件中添加配置节定义：
  <configSections>
    <section name="Giga.Transformer" type="Giga.Transformer.Configuration.TransformerConfigSection,Giga.Transformer"/>
  </configSections>
3. 根据项目的需要配置<Giga.Transformer>。主要是配置模版。具体参见Giga.Transformer.config文件。
4. 在代码中把配置读出来：
	TransformerConfigSection cfg = 。。。
5. 声明一个Transformer：
	var transformer = new Giga.Transformer.Transformer(cfg);
6. 如果是读取表格数据。比如类型为 TestRow。
	var entities = transformer.Load<TestRow>(filePath, "TestRowTemplate");
	其中，TestRowTemplate是你在配置中为TestRow数据读取而定义的模版名。
7. 如果是读取表单和嵌套数据，比如类型为 TestForm。
	var entity = transformer.LoadOne<TestForm>(filePath, "TestFormTemplate");
	其中，系统会自动根据配置把TestForm下的子表读出来。

----- 1.2 新功能 ---------------------------
1. Excel工作表复制功能
使用 ExcelUtils.CopyWorksheet(String srcFile, String srcSheet, String tgtFile, uint tgtPos) 函数复制工作表。


Thanks！
Shawn Xuehong Geng