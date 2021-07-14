package com.insolu.openoffice;

import java.io.File;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.List;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.regex.Pattern;

import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;

public class OpenOffice2PDF extends Thread {

    private static OpenOffice2PDF oop = new OpenOffice2PDF();

    private List<String> Office_Formats = new ArrayList<String>();

    private List<String> converterQueue = new CopyOnWriteArrayList<String>();

    private OfficeManager officeManager;

    private static final String[] OFFICE_POSTFIXS = { "doc", "docx", "xls",
            "xlsx", "ppt", "pptx" };

    private static final String PDF_POSTFIX = "pdf";

    // 设置任务执行超时时间， 分钟为单位
    private static final long TASK_EXECUTION_TIMEOUT = 1000 * 60 * 5L;

    // 设置任务队列超时时间，分钟为单位
    private static final long TASK_QUEUE_TIMEOUT = 1000 * 60 * 60 * 24L;

    private byte[] lock = new byte[0];

    private OpenOffice2PDF() {

    }

    public static OpenOffice2PDF getInstance() {
        return oop;
    }

    /**
     * 使Office2003-2007全部格式的文档(.doc|.docx|.xls|.xlsx|.ppt|.pptx) 转化为pdf文件
     *
     * @param inputFilePath
     *            源文件路径，如："e:/test.docx"
     * @param outputFilePath
     *            如果指定则按照指定方法，如果未指定（null）则按照源文件路径自动生成目标文件路径，如："e:/test_docx.pdf"
     * @return
     */
    private boolean converter(String inputFilePath, String outputFilePath) {
        boolean flag = false;
        long begin_time = new Date().getTime();
        File inputFile = new File(inputFilePath);
        if ((null != inputFilePath) && (inputFile.exists())) {
            Collections.addAll(Office_Formats, OFFICE_POSTFIXS);
            if (Office_Formats.contains(getPostfix(inputFilePath))) {
                try {
                    startService();
                    OfficeDocumentConverter converter = new OfficeDocumentConverter(
                            officeManager);
                    if (null == outputFilePath) {
                        outputFilePath = generateDefaultOutputFilePath(inputFilePath);
                    }
                    File outputFile = new File(outputFilePath);
                    converterFile(inputFile, outputFile, converter);
                    flag = true;
                } catch (Exception e) {
                    e.printStackTrace();
                } finally {
                    stopService();
                }
            }
        } else {
            System.out.println("con't find the resource");
        }

        long end_time = new Date().getTime();
        System.out.println("文件转换耗时：[" + (end_time - begin_time) + "]ms");
        return flag;
    }

    private void converterFile(File inputFile, File outputFile,
                               OfficeDocumentConverter converter) {

        if (!outputFile.getParentFile().exists()) {
            outputFile.getParentFile().mkdirs();
        }
        converter.convert(inputFile, outputFile);
    }

    public void add(String inputFile) {
        synchronized (converterQueue) {
            converterQueue.add(inputFile);
        }
    }

    /**
     *
     * @param converterQueue
     * @return
     */
    private String getTaskQueue(List<String> converterQueue) {
        synchronized (lock) {

            if (converterQueue.size() > 0) {
                return converterQueue.remove(0);
            } else {
                return null;
            }
        }
    }

    public void run() {
        while (true) {
            String taskInfo = getTaskQueue(converterQueue);
            if (null != taskInfo && taskInfo.trim().length() > 0) {
                converter(taskInfo, null);
            } else {
                try {
                    Thread.sleep(60000);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
            }
        }

    }

    private void startService() {
        DefaultOfficeManagerConfiguration config = new DefaultOfficeManagerConfiguration();
        String officeHome = getOfficeHome();
        config.setOfficeHome(officeHome);
        config.setTaskExecutionTimeout(TASK_EXECUTION_TIMEOUT);
        config.setTaskQueueTimeout(TASK_QUEUE_TIMEOUT);
        officeManager = config.buildOfficeManager();
        officeManager.start();
    }

    private void stopService() {
        if (officeManager != null) {
            officeManager.stop();
        }
    }

    /**
     * 根据操作系统的名称，获取OpenOffice.org 3的安装目录 如我的OpenOffice.org 3安装在：C:/Program
     * Files/OpenOffice.org 3
     */

    private String getOfficeHome() {
        String osName = System.getProperty("os.name");
        if (Pattern.matches("Linux.*", osName)) {
            return "/opt/openoffice.org3";
        } else if (Pattern.matches("Windows.*", osName)) {
            return "D:/OpenOffice4.1.3";
        }
        return null;
    }

    /**
     * 如果未设置输出文件路径则按照源文件路径和文件名生成输出文件地址。例，输入为 D:/fee.xlsx 则输出为D:/fee_xlsx.pdf
     */
    private String generateDefaultOutputFilePath(String inputFilePath) {
        String outputFilePath = (inputFilePath).replaceAll("."
                + getPostfix(inputFilePath), "_" + System.currentTimeMillis()
                + "." + PDF_POSTFIX);
        return outputFilePath;
    }

    /**
     * 获取inputFilePath的后缀名，如："e:/test.pptx"的后缀名为："pptx"
     */
    private String getPostfix(String inputFilePath) {
        String[] p = inputFilePath.split("\\.");
        if (p.length > 0) {// 判断文件有无扩展名
            // 比较文件扩展名
            return p[p.length - 1];
        } else {
            return null;
        }
    }
}
