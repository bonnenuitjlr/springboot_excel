package com.xingyu.execl.util;

import com.xingyu.execl.dto.response.User;
import com.xingyu.execl.util.execl.ExcelLog;
import com.xingyu.execl.util.execl.ExcelLogs;
import com.xingyu.execl.util.execl.ExcelUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;
import java.util.stream.Stream;

public class ExcelExportUtils {
    private static final int SHEET_SIZE_LIMIT = 60000;

    /**
     * 将execl导入成list
     * @param execl 需要导入的文件
     */
    public static void uploadExecl(MultipartFile execl) throws Exception{

        ExcelLogs excelLogs = new ExcelLogs();
        Collection<User> users = ExcelUtil.importExcel(User.class, execl.getInputStream(), "yyy-MM-dd", excelLogs, 0);
        System.out.println(users.toString());
        List<ExcelLog> errorLogList = excelLogs.getErrorLogList();
        for (ExcelLog log:errorLogList) {
            System.out.println(log.toString());
        }

    }

    /**
     * 将execl导入成list 自定义
     * @param execl 需要导入的文件
     */
    public static void uploadExecl2(MultipartFile execl) throws Exception{

        Map<String,Object> map = new HashMap<String,Object>();
        Map<String,Object> errorDetails = new HashMap<String,Object>();
        //1.得到上传的表
        Workbook workbook2 = WorkbookFactory.create(execl.getInputStream());
        //2、获取test工作表
        Sheet sheet2 = workbook2.getSheet("用户");
        //获取表的总行数
        int num = sheet2.getLastRowNum();
        //上传成功条数
        int successCount = 0;
        //上传失败条数
        int failCount = 0;
        //判断字段顺序 是否正确
        Row row0 = sheet2.getRow(0);
        Cell cell01 = row0.getCell(0);
        cell01.setCellType(CellType.STRING);
        Cell cell02 = row0.getCell(1);
        cell02.setCellType(CellType.STRING);
        Cell cell03 = row0.getCell(2);
        cell03.setCellType(CellType.STRING);

        for (int j = 1; j <= num; j++) {
            //这里new 一个对象，用来装填从页面上传的Excel数据，字段根据上传的excel决定
            User user= new User();
            boolean flag = true;
            Row row1 = sheet2.getRow(j);
            if(row1==null) {
                errorDetails.put("整行数据为空", "第"+(j+1)+"行");
                failCount += 1;
                continue;
            }
            //如果单元格中有数字或者其他格式的数据，则调用setCellType()转换为string类型
            Cell cell1 = row1.getCell(0);
            if(cell1!=null) {
                cell1.setCellType(CellType.STRING);
                user.setName(cell1.getStringCellValue());
            }else {
                errorDetails.put("姓名为空:第"+(j+1)+"行第1列","");
                flag = false;
            }
            //获取表中第i行，第2列的单元格
            Cell cell2 = row1.getCell(1);
            if(cell2!=null) {
                cell2.setCellType(CellType.STRING);
                if(!StringUtils.isEmpty(cell2.getStringCellValue())){
                    user.setSex(cell2.getStringCellValue().equals("男")?"1":"2");
                }else {
                    user.setSex("未定义");
                }
            }else {
                errorDetails.put("性别为空"+"第"+(j+1)+"行第2列","");
                flag = false;
            }
            //excel表的第i行，第3列的单元格
            Cell cell3 = row1.getCell(2);
            if(cell3!=null) {
                cell3.setCellType(CellType.STRING);
                user.setAge(cell3.getStringCellValue());
            }else {
                errorDetails.put("年龄为空:第"+(j+1)+"行第3列","");
                flag = false;
            }
            if(flag) {
                System.out.println(user.toString());
                successCount +=1;
            }else {
                failCount += 1;
            }

        }

    }

    /**
     * 将集合数据导出成excel字节xls格式
     *
     * @param <T>  类型
     * @param data 集合数据
     */
    public static <T> byte[] export2Byte(Collection<T> data) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        if (data == null || data.isEmpty()) {
            workbook.createSheet("无数据");
            workbook.write(out);
            return out.toByteArray();
        }
        T next = data.iterator().next();
        List<Field> list = sortedExcelFields(next.getClass());
        ExcelSheet excelSheet = next.getClass().getDeclaredAnnotation(ExcelSheet.class);
        String sheetNamePrefix = (excelSheet == null || excelSheet.value().isEmpty()) ? next.getClass().getName() : excelSheet.value();
        HSSFSheet sheet = null;
        int sheetCount = 0;
        int rownum = 0;
        for (T t : data) {
            if (rownum % SHEET_SIZE_LIMIT == 0) {
                if (sheetCount == 0) {
                    sheet = workbook.createSheet(sheetNamePrefix);
                }else{
                    sheetCount++;
                    sheet = workbook.createSheet(sheetNamePrefix + sheetCount);
                }
                writeTitle(sheet, list);
                rownum = 0;
            }
            rownum++;
            HSSFRow row = sheet.createRow(rownum);
            int columnindex = 0;
            for (Field f : list) {
                Object o = getProperty(t, f.getName());
                Cell cell = row.createCell(columnindex++, CellType.STRING);
                ExcelCell excelCell = f.getDeclaredAnnotation(ExcelCell.class);
                if (excelCell.valueAdaptor() != Adaptor.AdaptorDefault.class) {
                    try {
                        Adaptor adaptor = excelCell.valueAdaptor().newInstance();
                        o = adaptor.adaptor(o);
                        cell.setCellValue(get(o));
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                } else if (o != null && o instanceof Date) {
                    if (StringUtils.hasText(excelCell.pattern())) {
                        cell.setCellValue(DateUtil.formatDate(excelCell.pattern(), (Date) o));
                    } else {
                        cell.setCellValue(get(o));
                    }
                } else {
                    cell.setCellValue(get(o));
                }
            }
        }
        workbook.write(out);
        return out.toByteArray();
    }

    private static String get(Object v) {
        if (v == null) {
            return "";
        } else {
            return String.valueOf(v);
        }
    }

    private static void writeTitle(HSSFSheet sheet, List<Field> fields) {
        HSSFRow row = sheet.createRow(0);
        for (int i = 0; i < fields.size(); i++) {
            Field f = fields.get(i);
            Cell cell = row.createCell(i, CellType.STRING);
            ExcelCell excelCell = f.getDeclaredAnnotation(ExcelCell.class);
            cell.setCellValue(excelCell != null && !excelCell.value().isEmpty() ? excelCell.value() : f.getName());
        }
    }

    /**
     * 排序
     *
     * @param t   实例类
     * @param <T> 类型
     * @return 排序后的字段列表
     */
    private static <T> List<Field> sortedExcelFields(Class<T> t) {
        List<Field> fields = getExcelFields(t);
        fields.sort(Comparator.comparingInt(o -> o.getDeclaredAnnotation(ExcelCell.class).index()));
        return fields;
    }

    /**
     * excel字段列表
     *
     * @param t 实例
     * @return excel字段列表
     */
    private static List<Field> getExcelFields(Class<?> t) {
        List<Field> fields = new ArrayList<>();
        for (Class<?> clazz = t; clazz != Object.class && clazz != Class.class && clazz != Field.class; clazz = clazz.getSuperclass()) {
            try {
                Stream.of(clazz.getDeclaredFields()).
                        filter(field -> field.isAnnotationPresent(ExcelCell.class)).
                        forEach(fields::add);
            } catch (SecurityException ignore) {
            }
        }
        return fields;
    }

    private static <T> Object getProperty(T t, String propertyName) {
        try {
            PropertyDescriptor propertyDescriptor = new PropertyDescriptor(propertyName, t.getClass());
            Method method = propertyDescriptor.getReadMethod();
            if (!method.isAccessible()) {
                method.setAccessible(true);
            }
            return method.invoke(t);
        } catch (IntrospectionException | IllegalAccessException | IllegalArgumentException
                | InvocationTargetException e) {
            throw new RuntimeException("获取属性值失败" + e.getMessage(), e);
        }
    }

    @Retention(RetentionPolicy.RUNTIME)
    @Target(ElementType.FIELD)
    public @interface ExcelCell {
        /**
         * 顺序 default 100
         *
         * @return 顺序
         */
        int index() default 0;

        /**
         * 列名
         *
         * @return 列名
         */
        String value() default "";

        /**
         * 格式化模式date类型
         *
         * @return 格式化模式
         */
        String pattern() default "";

        /**
         * 格式转换
         */
        Class<? extends Adaptor> valueAdaptor() default Adaptor.AdaptorDefault.class;

        /**
         * 当值为null时要显示的值 default StringUtils.EMPTY
         *
         * @return defaultValue
         */
        String defaultValue() default "";

        /**
         * 用于验证
         *
         * @return valid
         */
       Valid valid() default @Valid();

        @Retention(RetentionPolicy.RUNTIME)
        @Target(ElementType.FIELD)
        @interface Valid {
            /**
             * 必须与in中String相符,目前仅支持String类型
             *
             * @return e.g. {"key","value"}
             */
            String[] in() default {};

            /**
             * 是否允许为空,用于验证数据 default true
             *
             * @return allowNull
             */
            boolean allowNull() default true;

            /**
             * Apply a "greater than" constraint to the named property
             *
             * @return gt
             */
            double gt() default Double.NaN;

            /**
             * Apply a "less than" constraint to the named property
             * @return lt
             */
            double lt() default Double.NaN;

            /**
             * Apply a "greater than or equal" constraint to the named property
             *
             * @return ge
             */
            double ge() default Double.NaN;

            /**
             * Apply a "less than or equal" constraint to the named property
             *
             * @return le
             */
            double le() default Double.NaN;
        }
    }

    public interface Adaptor<T> {

        Object adaptor(T t);

        class AdaptorDefault<T> implements Adaptor<T> {
            @Override
            public Object adaptor(T data) {
                return data;
            }
        }
    }


    @Retention(RetentionPolicy.RUNTIME)
    @Target(ElementType.TYPE)
    public @interface ExcelSheet {

        /**
         * sheet名
         *
         * @return sheet名
         */
        String value() default "";
    }

    @ExcelSheet("人民")
    private static class Man {

        @ExcelCell(value = "年龄", index = 2, valueAdaptor = AgeAdaptor.class)
        private int age;
        @ExcelCell(value = "姓名", index = 1)
        private String name;
        @ExcelCell(value = "日期", index = 3, pattern = "yyyyMMdd HH:mm:ss")
        private Date date;

        public int getAge() {
            return age;
        }

        public void setAge(int age) {
            this.age = age;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public Date getDate() {
            return date;
        }

        public void setDate(Date date) {
            this.date = date;
        }
    }

    public static class AgeAdaptor implements Adaptor<Integer> {
        public AgeAdaptor() {
        }

        @Override
        public Object adaptor(Integer integer) {
            if (integer > 18) {
                return "成年人";
            } else if (integer > 60) {
                return "老年人";
            }
            return "儿童";
        }
    }

    public static void main(String[] args) throws IOException {
        List<Man> ms = new ArrayList<>();
        for (int i = 0; i < 6; i++) {
            Man m = new Man();
            m.setAge(new Random().nextInt(80));
            m.setName("Alice " + m.getAge());
            m.setDate(new Date());
            ms.add(m);
        }
        byte[] bytes = export2Byte(ms);
        File f = new File("d:/a.xls");
        FileOutputStream outputStream = new FileOutputStream(f);
        outputStream.write(bytes);
        outputStream.close();
    }
}