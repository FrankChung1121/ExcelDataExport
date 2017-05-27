//
//  ViewController.m
//  ExcelDataExportDemo
//
//  Created by FrankChung on 17/5/27.
//  Copyright © 2017年 VMC. All rights reserved.
//

/**
 *  由于LibXL.framework库文件太大,不好上传,所以下载后拖进工程即可
 *  注意要设置bitcode为no,other linker flag也要改为-lstdc++
 */

#import "ViewController.h"
#import <LibXL/LibXL.h>

@interface ViewController () <UIDocumentInteractionControllerDelegate>

@property (nonatomic, strong) NSArray *nameArray;
@property (nonatomic, strong) NSArray *sexArray;
@property (nonatomic, strong) NSArray *ageArray;
@property (nonatomic, strong) NSArray *cityArray;
@property (nonatomic, strong) NSArray *schoolArray;
@property (nonatomic, strong) NSArray *phoneArray;
@property (nonatomic, strong) UIDocumentInteractionController *documentIc;

@end

@implementation ViewController

- (void)viewDidLoad {
    [super viewDidLoad];
    
    self.nameArray = @[@"张三", @"李四", @"王五", @"赵六", @"孙七"];
    self.sexArray = @[@"男", @"女", @"男", @"女", @"男"];
    self.ageArray = @[@"19", @"18", @"20", @"20", @"19"];
    self.cityArray = @[@"北京", @"上海", @"广州", @"深圳", @"杭州"];
    self.phoneArray = @[@"13888888888", @"13666666666", @"13999999999", @"13555555555", @"13777777777"];
    
    UIButton *exportBtn = [UIButton buttonWithType:UIButtonTypeCustom];
    exportBtn.frame = CGRectMake(100, 100, 200, 30);
    [exportBtn setTitle:@"导出excel表格数据" forState: UIControlStateNormal];
    [exportBtn setTitleColor:[UIColor brownColor] forState:UIControlStateNormal];
    [exportBtn addTarget:self action:@selector(exportBtnDidClick) forControlEvents:UIControlEventTouchUpInside];
    [self.view addSubview:exportBtn];
}

- (void)exportBtnDidClick {
    
    // 创建excel文件,表格的格式是xls,如果要创建xlsx表格,需要用xlCreateXMLBook()创建
    BookHandle book = xlCreateBook();
    
    // 创建sheet表格
    SheetHandle sheet = xlBookAddSheet(book, "Sheet1", NULL);
    
    /**
     *  设置表格的列宽
     *  参数1:数据要写入的表格
     *  参数2:从哪一列开始
     *  参数3:到哪一列结束
     *  参数4:具体的列宽
     *  参数5:数据要转换的格式,类型是FormatHandle,不清楚怎么定义的话可以直接写0,使用默认的
     *  参数6:隐藏属性,true
     */
    xlSheetSetCol(sheet, 4, 4, 15, 0, true);
    
    /**
     *  第一行的标题数据
     *  参数1:数据要写入的表格
     *  参数2:写入到哪一行
     *  参数3:写入到哪一列
     *  参数4:要写入的具体内容,注意是C字符串
     *  参数5:数据要转换的格式,类型是FormatHandle,不清楚怎么定义的话可以直接写0,使用默认的
     */
    xlSheetWriteStr(sheet, 1, 0, "姓名", 0);
    xlSheetWriteStr(sheet, 1, 1, "性别", 0);
    xlSheetWriteStr(sheet, 1, 2, "年龄", 0);
    xlSheetWriteStr(sheet, 1, 3, "城市", 0);
    xlSheetWriteStr(sheet, 1, 4, "电话", 0);
    
    // 从第二行开始写入数据,先把OC字符串转成C字符串
    for (int i = 0; i < self.nameArray.count; i++) {
        const char *name = [self.nameArray[i] cStringUsingEncoding:NSUTF8StringEncoding];
        xlSheetWriteStr(sheet, i + 2, 0, name, 0);
    }
    
    for (int i = 0; i < self.sexArray.count; i++) {
        const char *sex = [self.sexArray[i] cStringUsingEncoding:NSUTF8StringEncoding];
        xlSheetWriteStr(sheet, i + 2, 1, sex, 0);
    }
    
    for (int i = 0; i < self.ageArray.count; i++) {
        const char *age = [self.ageArray[i] cStringUsingEncoding:NSUTF8StringEncoding];
        xlSheetWriteStr(sheet, i + 2, 2, age, 0);
    }
    
    for (int i = 0; i < self.cityArray.count; i++) {
        const char *city = [self.cityArray[i] cStringUsingEncoding:NSUTF8StringEncoding];
        xlSheetWriteStr(sheet, i + 2, 3, city, 0);
    }
    
    for (int i = 0; i < self.phoneArray.count; i++) {
        const char *phone = [self.phoneArray[i] cStringUsingEncoding:NSUTF8StringEncoding];
        xlSheetWriteStr(sheet, i + 2, 4, phone, 0);
    }
    
    // 先写入沙盒
    NSString *documentPath = [NSSearchPathForDirectoriesInDomains(NSDocumentDirectory, NSUserDomainMask, YES)firstObject];
    NSString *fileName = [@"student" stringByAppendingString:@".xls"];
    NSString *filePath = [documentPath stringByAppendingPathComponent:fileName];
    
    // 保存表格
    xlBookSave(book, [filePath UTF8String]);
    
    // 最后要release表格
    xlBookRelease(book);
    
    // 调用safari分享功能将文件分享出去
    UIDocumentInteractionController *documentIc = [UIDocumentInteractionController interactionControllerWithURL:[NSURL fileURLWithPath:filePath]];
    
    // 记得要强引用UIDocumentInteractionController,否则控制器释放后再次点击分享程序会崩溃
    self.documentIc = documentIc;
    
    // 如果需要其他safari分享的更多交互,可以设置代理
    documentIc.delegate = self;
    
    // 设置分享显示的矩形框
    CGRect rect = CGRectMake(0, 0, 300, 300);
    [documentIc presentOpenInMenuFromRect:rect inView:self.view animated:YES];
    [documentIc presentPreviewAnimated:YES];
}


@end
