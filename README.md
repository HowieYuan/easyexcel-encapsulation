# easyexcel-method-encapsulation
easyexcel 项目地址 ：https://github.com/alibaba/easyexcel

#### 对 easyexcel 进行了方法的封装，可以做到一个函数完成简单的读取和导出

#### 目前 easyexcel 版本已经更新至 1.1.2-beta4

---

# 一. 依赖
首先是添加该项目的依赖，目前的版本是 1.0.2
```
<dependency>
	<groupId>com.alibaba</groupId>
	<artifactId>easyexcel</artifactId>
	<version>1.0.2</version>
</dependency>
```

# 二. 需要的类
![](https://upload-images.jianshu.io/upload_images/8807674-5fe0519ac2597f96.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

## 1. ExcelUtil
工具类，可以直接调用该工具类的方法完成 Excel 的读或者写

## 2. ExcelListener
监听类，可以根据需要，自定义处理获取到的数据
```
public class ExcelListener extends AnalysisEventListener {

    //自定义用于暂时存储data。
    //可以通过实例获取该值
    private List<Object> datas = new ArrayList<>();

    /**
     * 通过 AnalysisContext 对象还可以获取当前 sheet，当前行等数据
     */
    @Override
    public void invoke(Object object, AnalysisContext context) {
        //数据存储到list，供批量处理，或后续自己业务逻辑处理。
        datas.add(object);
        //根据自己业务做处理
        doSomething(object);
    }

    private void doSomething(Object object) {
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        /*
            datas.clear();
            解析结束销毁不用的资源
         */
    }

    public List<Object> getDatas() {
        return datas;
    }

    public void setDatas(List<Object> datas) {
        this.datas = datas;
    }
}
```

## 3. ExcelWriterFactroy
用于导出多个 sheet 的 Excel，通过多次调用 write 方法写入多个 sheet

## 4. ExcelException
捕获相关 Exception

# 三. 读取 Excel
读取 Excel 时只需要调用 ```ExcelUtil.readExcel()``` 方法
```
@RequestMapping(value = "readExcel", method = RequestMethod.POST)
public Object readExcel(MultipartFile excel) {
    return ExcelUtil.readExcel(excel, new ImportInfo());
}
```

其中 excel 是 MultipartFile 类型的文件对象，而 new ImportInfo() 是该 Excel 所映射的实体对象，需要继承 **BaseRowModel** 类，如：
```
public class ImportInfo extends BaseRowModel {
    @ExcelProperty(index = 0)
    private String name;

    @ExcelProperty(index = 1)
    private String age;

    @ExcelProperty(index = 2)
    private String email;

    /*
        作为 excel 的模型映射，需要 setter 方法
     */
    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getAge() {
        return age;
    }

    public void setAge(String age) {
        this.age = age;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }
}
```
作为映射实体类，通过 @ExcelProperty 注解与 index 变量可以标注成员变量所映射的列，同时不可缺少 setter 方法



# 四. 导出 Excel
### 1. 导出的 Excel 只拥有一个 sheet
只需要调用 ```ExcelUtil.writeExcelWithSheets()``` 方法：
```
@RequestMapping(value = "writeExcel", method = RequestMethod.GET)
public void writeExcel(HttpServletResponse response) throws IOException {
    List<ExportInfo> list = getList();
    String fileName = "一个 Excel 文件";
    String sheetName = "第一个 sheet";

    ExcelUtil.writeExcelWithSheets(response, list, fileName, sheetName, new ExportInfo());
    }
```
fileName，sheetName 分别是导出文件的文件名和 sheet 名，new ExportInfo() 为导出数据的映射实体对象，list 为导出数据。

对于映射实体类，可以根据需要通过 @ExcelProperty 注解自定义表头，当然同样需要继承 BaseRowModel 类，如：
```
public class ExportInfo extends BaseRowModel {
    @ExcelProperty(value = "姓名" ,index = 0)
    private String name;

    @ExcelProperty(value = "年龄",index = 1)
    private String age;

    @ExcelProperty(value = "邮箱",index = 2)
    private String email;

    @ExcelProperty(value = "地址",index = 3)
    private String address;
}
```
value 为列名，index 为列的序号

如果需要复杂一点，可以实现如下图的效果：
![](https://upload-images.jianshu.io/upload_images/8807674-5cb70346428fea93.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

对应的实体类写法如下：
```
public class MultiLineHeadExcelModel extends BaseRowModel {

    @ExcelProperty(value = {"表头1","表头1","表头31"},index = 0)
    private String p1;

    @ExcelProperty(value = {"表头1","表头1","表头32"},index = 1)
    private String p2;

    @ExcelProperty(value = {"表头3","表头3","表头3"},index = 2)
    private int p3;

    @ExcelProperty(value = {"表头4","表头4","表头4"},index = 3)
    private long p4;

    @ExcelProperty(value = {"表头5","表头51","表头52"},index = 4)
    private String p5;

    @ExcelProperty(value = {"表头6","表头61","表头611"},index = 5)
    private String p6;

    @ExcelProperty(value = {"表头6","表头61","表头612"},index = 6)
    private String p7;

    @ExcelProperty(value = {"表头6","表头62","表头621"},index = 7)
    private String p8;

    @ExcelProperty(value = {"表头6","表头62","表头622"},index = 8)
    private String p9;
}
```
### 2. 导出的 Excel 拥有多个 sheet
调用 ```ExcelUtil.writeExcelWithSheets()``` 处理第一个 sheet，之后调用 ```write()``` 方法依次处理之后的 sheet，最后使用 ```finish()``` 方法结束
```
public void writeExcelWithSheets(HttpServletResponse response) throws IOException {
    List<ExportInfo> list = getList();
    String fileName = "一个 Excel 文件";
    String sheetName1 = "第一个 sheet";
    String sheetName2 = "第二个 sheet";
    String sheetName3 = "第三个 sheet";

    ExcelUtil.writeExcelWithSheets(response, list, fileName, sheetName1, new ExportInfo())
                .write(list, sheetName2, new ExportInfo())
                .write(list, sheetName3, new ExportInfo())
                .finish();
}
```
write 方法的参数为当前 sheet 的 list 数据，当前 sheet 名以及对应的映射类

