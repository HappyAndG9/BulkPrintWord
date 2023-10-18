import cn.hutool.core.io.file.FileNameUtil;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.printing.PDFPrintable;
import org.apache.pdfbox.printing.Scaling;

import javax.print.PrintService;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.standard.Sides;
import java.awt.print.*;
import java.io.File;
import java.io.IOException;
import java.text.DecimalFormat;

/**
 * @Author 木月丶
 * @Description 主函数
 */

public class Print {
    //不输出日志信息
    static {
        java.util.logging.Logger
                .getLogger("org.apache.pdfbox").setLevel(java.util.logging.Level.OFF);
    }

    private static int pageCount;
    private static long time;
    private static final int printTimes = 1;  //打印多少份

    public static void main(String[] args) throws Exception {

        //设置打印机
        String printerName = "HP LaserJet P2015";

        //word文件夹的路径
        String wordFilePath = "F:\\Print_File\\Word";
        //存放生成的PDF文件的文件夹路径
        String PDFFilePath = "F:\\Print_File\\PDF";

        int odd = 0;    //页数为奇数的文档数
        int even = 0;   //页数为偶数的文档数

        //生成Word文件File[]
        File filesWord = new File(wordFilePath);
        File[] words = filesWord.listFiles();


        //将所有的word文档转化为PDF文件
        if(words != null && words.length != 0){
            for (int i = 0; i < words.length; i++) {
                //重命名
                String toFilePath = PDFFilePath + "\\" + FileNameUtil.mainName(words[i]) + ".pdf";
                //好看的分割线
                System.out.println("-----------↓--" + (i+1) + "--↓-----------");
                //word格式的文件转成pdf格式的文件
                wordToPDF(words[i].getAbsolutePath(),toFilePath);
                //到最后统计一下
                if (i == words.length-1){
                    System.out.println("-------------一共" + words.length + "个文件-------------");
                    //娱乐计时
                    System.out.println("---总耗时：" + time/1000.0 + "秒---");
                }
            }
        }

        //生成PDF文件的File[]
        File file = new File(PDFFilePath);
        int time = 1;       //循环次数  为了好看的输出而已
        int paper = 0,pages = 0;        //这个文件打印要用多少张纸    这个文件一共多少页
        DecimalFormat decimalFormat = new DecimalFormat("#.00");
        double amount = 0,totalAmount = 0;  //打印这个文件多少钱         打印全部文件一共多少钱

        //打印所有PDF文件
        for (File listFile : file.listFiles()) {
            for (int i = 0; i < printTimes; i++) {
                PDFPrint(listFile, printerName,time);
            }
            //获取这个文件的总页数
            pages = getPages(listFile);
            //通过总页数判断文件打印需要多少张纸，并统计奇数页文档和偶数页文档的个数
            if (pages % 2 == 0) {
                paper = pages/2;
                even++;
            } else {
                paper = (pages/2)+1;
                odd++;
            }
            //累加所有文件的总页数
            pageCount += pages;
            //单个文件打印费用
            amount = paper * 0.1 + pages *0.05;
            //累加所有文档打印的总费用
            totalAmount += amount;
            //循环次数+1 为了好看的输出
            time++;

            //输出单个文件的打印费用
//            System.out.println("-------------------------------");
//            System.out.println("【" + listFile.getName() + "】的打印费用为：" + amount);
//            if(time == listFile.length()) {
//                System.out.println("-------------------------------");
//            }

        }

        System.out.println("生成打印完毕！");
        System.out.println("一共 " + pageCount + " 页！");
        System.out.println("奇数页文档一共" + odd + "个。");
        System.out.println("偶数页文档一共" + even + "个。");
        System.out.println("一共：" + decimalFormat.format(totalAmount * printTimes)  + "元。");




    }

    /**
     * 打印PDF文件
     * @param file
     * @param printerName
     * @throws Exception
     */
    public static void PDFPrint(File file ,String printerName,int time) throws Exception {
        PDDocument document = null;     //声明PDF文件
        try {
            document = PDDocument.load(file);       //加载文件
            PrinterJob printJob = PrinterJob.getPrinterJob();   //创建打印事件
            printJob.setJobName(file.getName());    //设置事件名称

            // 查找并设置打印机
            if (printerName != null) {
                PrintService[] printServices = PrinterJob.lookupPrintServices();
                if(printServices == null || printServices.length == 0) {
                    System.out.print("打印失败，未找到可用打印机，请检查。");
                    return ;
                }

                PrintService printService = null;
                //
                //匹配指定打印机
                for (int i = 0;i < printServices.length; i++) {
                    if (printServices[i].getName().contains(printerName)) {         //包含子串即可，不需要全匹配
                        printService = printServices[i];
                        if(time == 1){
                            System.out.println("所用打印机：" + printService.getName());
                        }
                        break;
                    }
                }
                System.out.println(printService);
                //添加打印服务
                if(printService!=null){
                    printJob.setPrintService(printService);
                }else{
                    System.out.print("打印失败，未找到名称为" + printerName + "的打印机，请检查。");
                    return ;
                }
            }


            //设置纸张及缩放
            PDFPrintable pdfPrintable = new PDFPrintable(document, Scaling.ACTUAL_SIZE);    //实际大小
            //设置多页打印
            Book book = new Book();
            PageFormat pageFormat = new PageFormat();
            //设置打印方向
            pageFormat.setOrientation(PageFormat.PORTRAIT);     //纵向
            pageFormat.setPaper(getPaper());        //设置纸张
            book.append(pdfPrintable, pageFormat, document.getNumberOfPages());
            printJob.setPageable(book);
            printJob.setCopies(1);      //设置打印份数
            //添加打印属性
            HashPrintRequestAttributeSet pars = new HashPrintRequestAttributeSet();
            //设置单双页
            pars.add(Sides.DUPLEX);    //双面长边翻页
            //执行打印工作
            printJob.print(pars);
        }finally {
            //关闭资源
            if (document != null) {
                try {
                    document.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * 设置纸张和边距
     * @return
     */
    public static Paper getPaper() {
        Paper paper = new Paper();
        // 默认为A4纸张，对应像素宽和高分别为 595, 842
        int width = 595;
        int height = 842;
        // 设置边距，单位是像素，10mm边距，对应 28px
        int marginLeft = 10;
        int marginRight = 0;
        int marginTop = 10;
        int marginBottom = 0;
        paper.setSize(width, height);
        // 解决打印内容为空的问题
        paper.setImageableArea(marginLeft, marginRight, width - (marginLeft + marginRight), height - (marginTop + marginBottom));
        return paper;
    }

    /**
     * 将Word文件转为PDF文件
     * @param sFilePath
     * @param toFilePath
     */
    public static void wordToPDF(String sFilePath,String toFilePath) {
        System.out.println("启动 Word...");
        long start = System.currentTimeMillis();
        ActiveXComponent app = null;
        Dispatch doc = null;
        try {
            app = new ActiveXComponent("Word.Application");
            app.setProperty("Visible", new Variant(false));
            Dispatch docs = app.getProperty("Documents").toDispatch();
            doc = Dispatch.call(docs, "Open", sFilePath).toDispatch();
            System.out.println("打开文档:" + sFilePath);
            System.out.println("转换文档到 PDF:" + toFilePath);
            File toFile = new File(toFilePath);
            if (toFile.exists()) {
                toFile.delete();
            }
            Dispatch.call(doc, "SaveAs", toFilePath, // FileName
                    17);//17是pdf格式
            long end = System.currentTimeMillis();
            System.out.println("转换完成..用时：" + (end - start) + "ms.");
            time += (end -start);

        } catch (Exception e) {
            System.out.println("Error:文档转换失败：" + e.getMessage());
        } finally {
            Dispatch.call(doc, "Close", false);
            System.out.println("关闭文档");
            if (app != null)
                app.invoke("Quit", new Variant[]{});
        }
        // 释放资源
        ComThread.Release();
    }

    /**
     * 计算单个PDF文件的页数
     * @param pdf   文件
     * @return
     * @throws IOException
     */
    public static int getPages(File pdf) throws IOException {
        //加载PDF文件
        PDDocument load = PDDocument.load(pdf);
        //获取这个PDF的总页数
        int count = load.getPages().getCount();
        //关闭资源
        load.close();
        return count;
    }
}
