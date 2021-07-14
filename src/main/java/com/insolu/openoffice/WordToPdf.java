package com.insolu.openoffice;

import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;

import java.io.File;
import java.io.IOException;

/**
 * @program: openoffice_java
 * @description:
 * @author: xin.he
 * @create: 2021-07-14 13:49
 **/

public class WordToPdf {
    //private static final String wordInput = "D:\\insolu\\Doc\\05正方数字.docx";
    private static final String wordInput = "D:\\insolu\\Doc\\OCR服务部署手册.docx";
    private static final String pdfOutput = "D:\\insolu\\Doc\\0.outPut.pdf";

    public static void main(String[] args) {
        System.out.println("开始...");

        long startTime = System.currentTimeMillis(); //获取开始时间

        System.out.println(office2PDF(wordInput , pdfOutput));

        long endTime = System.currentTimeMillis(); //获取结束时间

        System.out.println("程序运行时间：" + (endTime - startTime) / 1000 + "s"); //输出程序运行时间

    }


    /**
     * 将Office文档转换为PDF. 运行该函数需要用到OpenOffice, OpenOffice下载地址为
     * http://www.openoffice.org/
     *
     * <pre>
     * 方法示例:
     * String sourcePath = "F:\\office\\source.doc";
     * String destFile = "F:\\pdf\\dest.pdf";
     * Converter.office2PDF(sourcePath, destFile);
     * </pre>
     *
     * @param sourceFile
     *            源文件, 绝对路径. 可以是Office2003-2007全部格式的文档, Office2010的没测试. 包括.doc,
     *            .docx, .xls, .xlsx, .ppt, .pptx等. 示例: F:\\office\\source.doc
     * @param destFile
     *            目标文件. 绝对路径. 示例: F:\\pdf\\dest.pdf
     * @return 操作成功与否的提示信息. 如果返回 -1, 表示找不到源文件, 或url.properties配置错误; 如果返回 0,
     *         则表示操作成功; 返回1, 则表示转换失败
     */
    public static int office2PDF(String sourceFile, String destFile) {
        try {
            File inputFile = new File(sourceFile);
            if (!inputFile.exists()) {
                return -1;// 找不到源文件, 则返回-1
            }

            // 如果目标路径不存在, 则新建该路径
            File outputFile = new File(destFile);
            if (!outputFile.getParentFile().exists()) {
                outputFile.getParentFile().mkdirs();
            }

            // connect to an OpenOffice.org instance running on port 8100
            // 连接到在端口 8100 上运行的 OpenOffice.org 实例
            OpenOfficeConnection connection = new SocketOpenOfficeConnection(
                    "127.0.0.1", 8100);
            connection.connect();

            // convert
            DocumentConverter converter = new OpenOfficeDocumentConverter(
                    connection);

             converter.convert(inputFile, outputFile);

            // close the connection
            connection.disconnect();

            return 0;
        } catch (IOException e) {
            e.printStackTrace();
        }

        return 1;
    }
}
