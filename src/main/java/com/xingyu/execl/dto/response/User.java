package com.xingyu.execl.dto.response;

import com.xingyu.execl.util.ExcelExportUtils;
import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * @author yangxingyu
 * @date 2019/12/23
 * @description
 */
@Data
@AllArgsConstructor
@ExcelExportUtils.ExcelSheet("用户")
public class User {

    public User(){}

    @ExcelExportUtils.ExcelCell(value = "姓名", index = 1)
    private String name;

    @ExcelExportUtils.ExcelCell( valid = @ExcelExportUtils.ExcelCell.Valid(allowNull = false,in = {"男","女"}), value = "性别", index = 2, valueAdaptor = UserAdaptor.class)
    private String sex;

    @ExcelExportUtils.ExcelCell(value = "年龄", index = 3)
    private String age;

}
