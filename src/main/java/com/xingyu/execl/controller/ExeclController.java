package com.xingyu.execl.controller;

import com.xingyu.execl.dto.response.User;
import com.xingyu.execl.util.ExcelExportUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Base64;
import java.util.List;

/**
 * @author yangxingyu
 * @date 2019/12/23
 * @description
 */
@RestController
public class ExeclController {

    @RequestMapping(value = "/getExcel", method = RequestMethod.GET)
    public String getExecl(HttpServletResponse response) {
        String result = "";
        List<User> list = new ArrayList<User>();
        User user1 = new User("text1", "1", "22");
        User user2 = new User("text2", "2", "23");
        User user3 = new User("text3", "1", "24");
        list.add(user1);
        list.add(user2);
        list.add(user3);
        try {
            //如果前端需要 字节数组 直接返回这个
            byte[] bytes = ExcelExportUtils.export2Byte(list);

            //如果前端需要 字符串 直接返回这个
            result = Base64.getEncoder().encodeToString(bytes);

            //如果需要 下载  直接返回这个
            String fileName = "fileName";
            response.reset();
            response.setContentType("application/vnd.ms-excel;charset=utf-8");
            response.setHeader("Content-Disposition", "attachment;filename=" + fileName+".xls");
            OutputStream outputStream = response.getOutputStream();
            outputStream.write(bytes);
            outputStream.flush();
            outputStream.close();

        } catch (Exception e) {

        }


        return "getExcel";
    }

    @RequestMapping(value = "/uploadExecl",method = RequestMethod.POST)
    public boolean uploadExecl(@RequestParam(value="excel") MultipartFile execl){
        try {
            ExcelExportUtils.uploadExecl(execl);
        }catch (Exception e){
            return false;
        }
        return true;
    }

}
