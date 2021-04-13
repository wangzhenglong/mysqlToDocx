
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import java.math.BigInteger;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.time.Duration;
import java.time.LocalDateTime;
import java.util.HashMap;
import java.util.LinkedList;

/**
 * @author dragonKj
 * @Description TODO
 * @createTime 2021/4/6  9:56
 */
public class Main {
    private static WordprocessingMLPackage wordMLPackage;
    private static ObjectFactory factory=Context.getWmlObjectFactory();
    //需要导出的数据库名称
    private static final String DATABASENAME="rec";

    public static void main(String[] args) {
        try {
            LocalDateTime start= LocalDateTime.now();;
            wordMLPackage = WordprocessingMLPackage.createPackage();
            LinkedList<HashMap<String,String>> linkedList=getTableName(DATABASENAME);
            for (HashMap<String,String> hashMap:linkedList) {
                addTable(wordMLPackage,hashMap);
            }

            wordMLPackage.save(new java.io.File("src/main/"+DATABASENAME+".docx"));
            LocalDateTime end=LocalDateTime.now();

            Duration duration = Duration.between(start,end);
            long millis=duration.toMillis();
            System.out.println("导出用时："+millis+"毫秒");
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    private static void addTable(WordprocessingMLPackage wordMLPackage,HashMap<String,String> hashMap) throws Exception {

        //表title
        String TABLE_NAME=hashMap.get("TABLE_NAME");
        String TABLE_COMMENT=hashMap.get("TABLE_COMMENT");
        String text="";
        if("".equals(TABLE_COMMENT)){
            text= TABLE_NAME;
        }else {
            text= TABLE_NAME+"("+TABLE_COMMENT+")";
        }
        wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1",text);
        //新建表格
        Tbl table = factory.createTbl();
        //新增表头
        addTableHead(table);
        //新增表行
        LinkedList<HashMap<String,String>> linkedList=getCOLUMNS(TABLE_NAME);
        int num=0;
        for (HashMap<String,String> hashMap1:linkedList) {
            num++;
            addTR(table,hashMap1,num);
        }

        //表格加边框
        addBorders(table);
        wordMLPackage.getMainDocumentPart().addObject(table);
        wordMLPackage.getMainDocumentPart().addParagraphOfText("");

    }


    /**
     * 表格加边框
     *
     * @param table
     */
    private static void addBorders(Tbl table) {
        table.setTblPr(new TblPr());
        CTBorder border = new CTBorder();
        border.setColor("auto");
        border.setSz(new BigInteger("4"));
        border.setSpace(new BigInteger("0"));
        border.setVal(STBorder.SINGLE);

        TblBorders borders = new TblBorders();
        borders.setBottom(border);
        borders.setLeft(border);
        borders.setRight(border);
        borders.setTop(border);
        borders.setInsideH(border);
        borders.setInsideV(border);
        table.getTblPr().setTblBorders(borders);
    }

    /**
     * 本方法创建单元格, 添加样式后添加到表格行中
     * <p>
     * 表格行，内容，是否加粗，字体大小,表格宽度
     */
    private static void addStyledTableCell(Tr tableRow, String content,
                                           boolean bold, String fontSize, int width) {
        Tc tableCell = factory.createTc();
        //设置表格样式
        addStyling(tableCell, content, bold, fontSize);
        if (width > 0) {
            //设置表格宽度
            setCellWidth(tableCell, width);
        }
        tableRow.getContent().add(tableCell);


    }

    /**
     * 这里我们添加实际的样式信息, 首先创建一个段落, 然后创建以单元格内容作为值的文本对象;
     * 第三步, 创建一个被称为运行块的对象, 它是一块或多块拥有共同属性的文本的容器, 并将文本对象添加
     * 到其中. 随后我们将运行块R添加到段落内容中.
     * 直到现在我们所做的还没有添加任何样式, 为了达到目标, 我们创建运行块属性对象并给它添加各种样式.
     * 这些运行块的属性随后被添加到运行块. 最后段落被添加到表格的单元格中.
     */
    private static void addStyling(Tc tableCell, String content, boolean bold, String fontSize) {
        P paragraph = factory.createP();

        Text text = factory.createText();
        text.setValue(content);

        R run = factory.createR();
        run.getContent().add(text);

        paragraph.getContent().add(run);

        RPr runProperties = factory.createRPr();
        if (bold) {
            addBoldStyle(runProperties);
        }

        if (fontSize != null && !fontSize.isEmpty()) {
            setFontSize(runProperties, fontSize);
        }

        run.setRPr(runProperties);

        tableCell.getContent().add(paragraph);
    }

    /**
     * 本方法为可运行块添加字体大小信息. 首先创建一个"半点"尺码对象, 然后设置fontSize
     * 参数作为该对象的值, 最后我们分别设置sz和szCs的字体大小.
     * Finally we'll set the non-complex and complex script font sizes, sz and szCs respectively.
     */
    private static void setFontSize(RPr runProperties, String fontSize) {
        HpsMeasure size = new HpsMeasure();
        size.setVal(new BigInteger(fontSize));
        runProperties.setSz(size);
        runProperties.setSzCs(size);
    }

    /**
     * 本方法给可运行块属性添加粗体属性. BooleanDefaultTrue是设置b属性的Docx4j对象, 严格
     * 来说我们不需要将值设置为true, 因为这是它的默认值.
     */
    private static void addBoldStyle(RPr runProperties) {
        BooleanDefaultTrue b = new BooleanDefaultTrue();
        b.setVal(true);
        runProperties.setB(b);
    }


    /**
     * 本方法创建一个单元格属性集对象和一个表格宽度对象. 将给定的宽度设置到宽度对象然后将其添加到
     * 属性集对象. 最后将属性集对象设置到单元格中.
     */
    private static void setCellWidth(Tc tableCell, int width) {
        TcPr tableCellProperties = new TcPr();
        TblWidth tableWidth = new TblWidth();
        tableWidth.setW(BigInteger.valueOf(width));
        tableCellProperties.setTcW(tableWidth);
        tableCell.setTcPr(tableCellProperties);

    }
    /**
     * 新建表头
     */
    private static void addTableHead(Tbl table){
        //新建表格表头
        Tr tableRow = factory.createTr();
        //表格表头填充内容
        //表格行，内容，是否加粗，字体大小,表格宽度
        addStyledTableCell(tableRow, "序号", true, "20", 700);
        addStyledTableCell(tableRow, "字段名称", true, "20", 2000);
        addStyledTableCell(tableRow, "数据类型", true, "20", 1200);
        addStyledTableCell(tableRow, "是否主键", true, "20", 1200);
        addStyledTableCell(tableRow, "可以为空", true, "20", 1200);
        addStyledTableCell(tableRow, "字段注释", true, "20", 3000);
        table.getContent().add(tableRow);
    }

    /**
     * 新建表行
     */
    private static void addTR(Tbl table,HashMap<String,String> hashMap,int num){
        //新建表格表行
        Tr tableRow = factory.createTr();
        //表格表行填充内容
        //表格行，内容，是否加粗，字体大小,表格宽度
        String COLUMN_NAME=hashMap.get("COLUMN_NAME");
        String COLUMN_TYPE=hashMap.get("COLUMN_TYPE");
        String COLUMN_COMMENT=hashMap.get("COLUMN_COMMENT");
        String column_key=hashMap.get("column_key");
        String IS_NULLABLE=hashMap.get("IS_NULLABLE");
        addStyledTableCell(tableRow, String.valueOf(num), false, "20", 700);
        addStyledTableCell(tableRow, COLUMN_NAME, false, "20", 2000);
        addStyledTableCell(tableRow, COLUMN_TYPE, false, "20", 1200);
        addStyledTableCell(tableRow, column_key, false, "20", 1200);
        addStyledTableCell(tableRow, IS_NULLABLE, false, "20", 1200);
        addStyledTableCell(tableRow, COLUMN_COMMENT, false, "20", 3000);
        table.getContent().add(tableRow);
    }

    /**
     * 获取表名和表注释
     * @param dataBaseName
     * @return
     * @throws Exception
     */
    private static LinkedList<HashMap<String,String>> getTableName(String dataBaseName) throws Exception{
        String driver="com.mysql.jdbc.Driver";
        String url="jdbc:mysql://192.168.7.60:3307/oss?useUnicode=true&amp;characterEncoding=utf-8&amp;autoReconnect=true&amp;zeroDateTimeBehavior=convertToNull";
        String user="root";
        String password="RHtj69*64admin%nimda46*96jtHR";

        LinkedList<HashMap<String,String>> linkedList=new LinkedList<HashMap<String, String>>();
        Connection con=null;
        PreparedStatement pre=null;
        ResultSet rs=null;
       try {
           Class.forName(driver);
           con= DriverManager.getConnection(url,user,password);
           String sql="SELECT\n" +
                   "\tTABLE_NAME AS TABLE_NAME,\n" +
                   "\tTABLE_COMMENT AS TABLE_COMMENT \n" +
                   "FROM\n" +
                   "\tINFORMATION_SCHEMA.TABLES \n" +
                   "WHERE\n" +
                   "\tTABLE_SCHEMA = ?\n";
           pre=con.prepareStatement(sql);
           pre.setString(1,dataBaseName);
           rs=pre.executeQuery();
           while(rs.next()){
               HashMap<String,String> map=new HashMap<String, String>(50);
               String TABLE_NAME=rs.getString(1);
               String TABLE_COMMENT=rs.getString(2);
               map.put("TABLE_NAME",TABLE_NAME);
               map.put("TABLE_COMMENT",TABLE_COMMENT);
               linkedList.add(map);
           }

       }catch (Exception e){
           new RuntimeException(e.getMessage());

       }finally {
           if(pre != null) pre.close();
           if(con != null) con.close();
           return linkedList;
       }
    }


    /**
     * 获取字段名和字段注释
     * @param tableName
     * @return
     * @throws Exception
     */
    private static LinkedList<HashMap<String,String>> getCOLUMNS(String tableName) throws Exception{
        String driver="com.mysql.jdbc.Driver";
        String url="jdbc:mysql://192.168.7.60:3307/oss?useUnicode=true&amp;characterEncoding=utf-8&amp;autoReconnect=true&amp;zeroDateTimeBehavior=convertToNull";
        String user="root";
        String password="RHtj69*64admin%nimda46*96jtHR";

        LinkedList<HashMap<String,String>> linkedList=new LinkedList<HashMap<String, String>>();
        Connection con=null;
        PreparedStatement pre=null;
        ResultSet rs=null;
        try {
            Class.forName(driver);
            con= DriverManager.getConnection(url,user,password);
            String sql="SELECT\n" +
                    "\tCOLUMN_NAME ,\n" +
                    "\tCOLUMN_TYPE ,\n" +
                    "\tCOLUMN_COMMENT ,\n" +
                    "\tIF(column_key='PRI','1','0') AS 'column_key',\n" +
                    "  IF(IS_NULLABLE='NO','0','1') AS 'IS_NULLABLE'\n" +
                    "FROM\n" +
                    "\tINFORMATION_SCHEMA.COLUMNS \n" +
                    "WHERE\n" +
                    "\tTABLE_SCHEMA = ? \n" +
                    "\tAND TABLE_NAME = ?";
            pre=con.prepareStatement(sql);
            pre.setString(1,DATABASENAME);
            pre.setString(2,tableName);
            rs=pre.executeQuery();
            while(rs.next()){
                HashMap<String,String> map=new HashMap<String, String>(50);
                String COLUMN_NAME=rs.getString(1);
                String COLUMN_TYPE=rs.getString(2);
                String COLUMN_COMMENT=rs.getString(3);
                String column_key=rs.getString(4);
                String IS_NULLABLE=rs.getString(5);
                map.put("COLUMN_NAME",COLUMN_NAME);
                map.put("COLUMN_TYPE",COLUMN_TYPE);
                map.put("COLUMN_COMMENT",COLUMN_COMMENT);
                if("1".equals(column_key)){
                    column_key="是";
                }else if("0".equals(column_key)){
                    column_key="否";
                }
                if("1".equals(IS_NULLABLE)){
                    IS_NULLABLE="是";
                }else if("0".equals(IS_NULLABLE)){
                    IS_NULLABLE="否";
                }
                map.put("column_key",column_key);
                map.put("IS_NULLABLE",IS_NULLABLE);
                linkedList.add(map);
            }

        }catch (Exception e){
            new RuntimeException(e.getMessage());

        }finally {
            if(pre != null) pre.close();
            if(con != null) con.close();
            return linkedList;
        }
    }

}
