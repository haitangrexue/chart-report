package com.zsir.chartreport.template;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.lang.Dict;
import cn.hutool.core.util.CharsetUtil;
import cn.hutool.extra.template.Template;
import cn.hutool.extra.template.TemplateConfig;
import cn.hutool.extra.template.TemplateEngine;
import cn.hutool.extra.template.TemplateUtil;
import org.junit.jupiter.api.Test;

/**
 * @ClassName Template
 * @description:
 * @author: zjj
 * @create: 2025-07-12 10:00
 **/
public class TemplateTest {

    @Test
    public void create() {
        // 创建引擎，指定模板文件在classpath的templates目录下
        TemplateEngine engine = TemplateUtil.createEngine(
                new TemplateConfig("templates", TemplateConfig.ResourceMode.CLASSPATH)
        );

        // 获取模板，注意模板文件在templates目录下，所以直接写文件名
        Template template = engine.getTemplate("report.ftl");

        // 构造填充数据
        Dict dict = Dict.create();
        dict.set("chart", "图表base64编码");

        // 模板填充后返回内容
        String content = template.render(dict);

        FileUtil.writeString(content, "d:/report.doc", CharsetUtil.CHARSET_UTF_8);
    }
}
