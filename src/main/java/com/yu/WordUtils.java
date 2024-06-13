package com.yu;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.file.FileReader;
import cn.hutool.poi.word.Word07Writer;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;

import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.util.*;
import java.util.List;

public class WordUtils {

    static Word07Writer writer = new Word07Writer();
    static final Scanner sc = new Scanner(System.in);


    // 排除的文件后缀
    static Set<String> excludes = new HashSet<>();
    static Set<String> exc = Set.of("svg","jpg","png","gif","docx","jar");
    // 排除的文件夹名称
    static Set<String> excludesDir = new HashSet<>();
    static Set<String> excDir = Set.of(".idea",".vscode",".git");
    // 排除的绝对路径
    static Set<String> excludesFull = new HashSet<>();
    static Set<String> excFull = Set.of("D:\\Work\\DormitoryManagementSystem\\vue3-element-admin\\src\\views\\demo\\table",
            "D:\\Work\\DormitoryManagementSystem\\vue3-element-admin\\src\\views\\demo\\multi-level");


    private static void writerFile(File file){
        FileReader fileReader = new FileReader(file);
        String title = "=======" + file.getAbsolutePath() + " ========";
        int count = 1;
        writer.addText(ParagraphAlignment.CENTER,new Font("Consolas", Font.PLAIN, 10), title);
        List<String> s = fileReader.readLines();
        for (String result : s) {
            writer.addText(new Font("Consolas", Font.PLAIN, 7)
                    ,"%3d.%s".formatted(count++,result));
        }
    }

    public static void exportCodeFiles(String directory,String outputPath){
        writer.setDestFile(FileUtil.file(outputPath));
        File[] ls = FileUtil.ls(directory);
        for (File file : ls) {
            if (excludesFull.contains(file.getAbsolutePath())){
                System.out.println("跳过文件："+file.getAbsolutePath());
                continue;
            }
            if (file.isFile()){
                if (excludes.contains(FileUtil.extName(file))) {
                    System.out.println("跳过文件："+file.getAbsolutePath());
                    continue;
                }
                writerFile(file);
            }else if (file.isDirectory()){
                if (excDir.contains(file.getName())){
                    System.out.println("跳过文件夹："+file.getAbsolutePath());
                    continue;
                }
                exportCodeFiles(file.getAbsolutePath(),outputPath);
            }
        }
    }

    public static void export() throws IOException {
        excludesDir.addAll(excDir);
        excludes.addAll(exc);
        excludesFull.addAll(excFull);
        System.out.println("请输入项目路径数量：");
        int num = sc.nextInt();
        ArrayList<String> directory = new ArrayList<>();
        for (int i = 0; i < num; i++) {
            System.out.println("请输入项目路径：");
            directory.add(sc.next());
        }
        System.out.println("请输入导出路径：(如：D:\\\\c.docx注意后缀为.docx)");
        String outputPath = sc.next();
        System.out.println("是否有需要排除的文件夹名？n则立即导出，是文件名(默认排除.git,.idea,.vscode)");
        inputExcludes(excludesDir);
        System.out.println("是否有需要排除的后缀？默认有(svg,jpg,png,gif,docx,jar)");
        System.out.println("n则立即导出");
        inputExcludes(excludes);
        System.out.println("是否有额外需要排除的文件/文件夹？n则立即导出，输入绝对路径");
        inputExcludes(excludes);
        for (String s : directory) {
            exportCodeFiles(s,outputPath);
        }
        writer.flush();
        writer.close();
        System.out.println("导出完成！");
        System.out.println("打开文件？");
        openFile(outputPath);
    }

    private static void openFile(String outputPath) throws IOException {
        Runtime runtime = Runtime.getRuntime();
        runtime.exec("rundll32 url.dll FileProtocolHandler "+ outputPath);
    }

    private static void inputExcludes(Set<String> excludes) {
        String excFull = sc.next();
        System.out.println("输入-1为终止");
        if (!"n".equals(excFull)){
            String isExcludes = sc.next();
            while (!"-1".equals(isExcludes)){
                excludes.add(isExcludes);
                isExcludes = sc.next();
            }
        }
    }
}
