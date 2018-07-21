# easyexcel-method-encapsulation
easyexcel 项目地址 ：https://github.com/alibaba/easyexcel

#### 对 easyexcel 进行了方法的封装，可以做到一个函数完成简单的读取和导出

#### 原项目目前仍存在一些BUG：
- ~~XLSX 类型的 EXCEL 在读取的时候，序号为 1 的 sheet 为最后一个 sheet（即 sheet 的顺序为倒序）；~~   
  ~~XLS 类型的 EXCEL 在读取的时候，序号为 1 的 sheet 为第一个 sheet（即 sheet 的顺序为顺序）；~~
- ~~将导出类型为 XLSX 的 excel 导入时，会报错~~  
~~而将导出类型为 XLS 的 excel 导入时，则不会~~


#### 目前 easyexcel 版本已经更新至 1.0.2，修复了一些 BUG
