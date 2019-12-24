package com.xingyu.execl.dto.response;

import com.xingyu.execl.util.ExcelExportUtils;

/**
 * @author yangxingyu
 * @date 2019/12/23
 * @description
 */
public class UserAdaptor implements ExcelExportUtils.Adaptor<String>{

    public UserAdaptor() {
    }

    @Override
    public Object adaptor(String type) {
        if ("1".equals(type)) {
            return "男";
        } else if ("2".equals(type)) {
            return "女";
        }
        return "";
    }
}
