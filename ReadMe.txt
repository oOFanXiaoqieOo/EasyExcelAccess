easyexcel-access
1.简介（1.0-SNAPSHOT）
使用easyexcel导出excel数据
静态工具类方法，存储一次历史读取结果，如果重读该数据，直接返回历史数据


2.使用说明
静态方法1：
----->使用List<Object> readExcelData(String filePath, int sheetNo)函数返回内容数据
输入参数为文件路径和表格号；全局输入参数为EasyExcelUtil.defaultHeadLineNum内容数据开始行号（0为第一行）

静态方法2：
----->使用ExcelListener readExcel(String filePath, int sheetNo)函数返回标题数据和内容数据  (ExcelListener为包内自定义类)  
输入参数同方法1
(List<Object>) ExcelListener.getDatas()为内容数据（同方法1返回数据）
(List<Object>) ExcelListener.getHead()为标题数据

方法3：使用数据挡板lMask（List<Integer>）
对excel中每一行执行一次挡板操作，lMask中的数据为excel列输出顺序（可重复、越界为空）
----->a.先设置lMask（重置重读标识，历史数据会失效）
----->b.再使用方法1或方法2读取数据，返回值会根据lMask变化,返回值形为List<Object>

方法4：使用数据挡板mMask（Map<String, Integer>）
对excel中每一行执行一次挡板操作，mMask中的数据为excel根据标题列的筛选数据（不可重复、越界为空、后匹配数据生效）
----->a.先设置mMask（重置重读标识，历史数据会失效）
----->b.再使用方法1或方法2读取数据，返回值会根据mMask变化,返回值形为List<Map<String,Object>>

方法5：使用数据挡板dMask（List<EasyExcelMask>）  (EasyExcelMask为包内自定义类)
对excel中每一行执行一次挡板操作，dMask可通过xml文件加载（支持精确匹配，范围匹配，关键字替换操作） 精确匹配优先于范围匹配，范围匹配先匹配数据生效
----->a.调用EasyExcelMask.LoadMask(path,"nodename")方法加载dMask数据
----->b.再使用方法1或方法2读取数据，返回值会根据dMask变化,返回值形为List<Map<String,Object>>

xml文件形如
<MaskList>
	<nodename>
		<Element MaskName="名字" FakeName="Name" Type="int" >AA,AB,AC</Element>
		<Element MaskName="年龄" FakeName="Age"  Type="Integer" >AA</Element>
		<Element FakeName="Sex" >AA,AB,AC</Element>
		<Element >BB,CC,DD</Element>
	</nodename>
</MaskList>

 *精确匹配数据		maskStr=名字(第一个property属性,MaskName优先选择)  	
 *关键字替换值		fakemask=Name(第二个property属性,FakeName优先选择)
 *数据类型			typeStrtypeStr=int(Type属性，多属性且名字不规范的情况下，Type属性不要排在前两个，有可能会影响maskStr和fakemask)（导入数据库使用）
 *范围匹配数据		matchStr=AA,AB,AC(text根据','转为list)		

 
 
升级.简介（1.1）
xml配置增加List属性（只能通过xml文件加载）
<Element MaskName="名字" FakeName="Name" Type="int" List="DList"></Element>
根据excel上下文组装一次list数据，组装规则，非List列都为空或者与上一次数据相同
excel同一行中若存在相同List属性值，则打包为<Map>中；非同一行List属性值打包为List<Map>
例：
xml文件：
<MaskList>
	<nodelistname>
		<Element MaskName="名字" FakeName="Name" ></Element>
		<Element MaskName="年龄" FakeName="Age" ></Element>
		<Element MaskName="性别" FakeName="Sex" ></Element>
		<Element MaskName="课程1" FakeName="Course1"  List="CourseList1" ></Element>
		<Element MaskName="课程2" FakeName="Course2"  List="CourseList1" ></Element>
		<Element MaskName="课程3" FakeName="Course3"  List="CourseList2" ></Element>
	</nodelistname>
</MaskList>

excel数据：
名字      年龄      性别      课程1     课程2     课程3
小明       3        1         A          B          C
                              AA         BB         CC
小小       2        2         微积分     天体物理   格斗
                              针织       C语言      演员的自我修养

组装数据：
{Age=3, 
CourseList1=[
	{Course2=B, Course1=A}, 
	{Course2=BB, Course1=AA}
],
Name=小明, 
Sex=1, 
CourseList2=[
	{Course3=C}, 
	{Course3=CC}
]},
{Age=2, CourseList1=[{Course2=天体物理, Course1=微积分}, {Course2=C语言, Course1=针织}], Name=小小, Sex=2, CourseList2=[{Course3=格斗}, {Course3=演员的自我修养}]}


示例代码：
	String filePath = "E://HONG/1.xlsx";
	String path = "E://HONG/1.xml";         //从配置文件中加载挡板组装配置数据
	EasyExcelMask load=new EasyExcelMask();
	List<EasyExcelMask> dMask =load.LoadMask(path,"nodelistname");//<-------------------加载dMask数据
	EasyExcelUtil.setdMask(dMask);
    List<Object> objects=EasyExcelUtil.readExcelData(filePath, 3);//<-------------------读取excel数据
    objects.forEach(System.out::println);