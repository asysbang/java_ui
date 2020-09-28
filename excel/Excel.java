package excel;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Stack;

import jxl.Cell;
import jxl.Workbook;
import jxl.format.CellFormat;
import jxl.write.*;

//  第8列文字描述也需要更改
//  case id 也需要对应的修改

public class Excel {

	private WritableSheet mTargetSheet;

	private WritableWorkbook mWorkbook;

	public static void main(String[] args) throws Exception {
		Excel excel = new Excel();
		excel.runTask();
	}

	private List<String> makeMultiContents(String contents) {
		System.out.println("=====原始数据 ： \n" + contents);
		List<String> res = new ArrayList<String>();
		List<String> contentsList = new ArrayList<String>();
		String[] lines = contents.split("语音录入");// 选择结果
		for (String str : lines) {
			// System.out.println("===str==" + str);
			String[] lines2 = str.split("选择结果");
			for (String str2 : lines2) {
				// System.out.println("===str2==" + str2);
				contentsList.add(str2);
			}
		}
		for (String str : contentsList) {
			System.out.println("===str ==" + str);
		}
		System.out.println("===size ==" + contentsList.size());
		stack.clear();
		result.clear();
		int size = contentsList.size();
		if(size <3) {
			return res;
		}
		f(shu, size - 2, 0);// 第一个都是语音输入,而且本身占位符个数就是size-1，所以这里需要-2
		for (int x = 0; x < result.size(); x++) {
			String[] words = result.get(x).split(",");
			System.out.println("====words.length="+words.length);
			System.out.println("===size=="+size);
			StringBuilder sb = new StringBuilder();
			for (int i = 0; i < size; i++) {
				sb.append(contentsList.get(i));
				if(i==size-1) {
					break;
				}
				if(i==0) {
					sb.append("语音录入");
				} else {
					sb.append(words[i-1]);
				}
			}
			res.add(sb.toString());
			System.out.println("=====改装数据(" + x + ") ： \n" + sb.toString());
		}

//		for (int x = 0; x < lines.length - 1; x++) {
//			StringBuilder sb = new StringBuilder();
//			for (int i = 0; i < lines.length; i++) {
//				sb.append(lines[i]);
//				if (i < lines.length - 1) {
//					if (i == x) {
//						sb.append("手工输入");
//					} else {
//						sb.append("语音录入");
//					}
//				}
//			}
//			res.add(sb.toString());
//			// System.out.println("=====改装数据(" + x + ") ： \n" + sb.toString());
//		}
		return res;
	}

	private void addAtRowIndex(int index) throws Exception {
		Cell cell = mTargetSheet.getCell(7, index);
		String contents = cell.getContents();
//		System.out.println("===" + index + "===\n" + contents);
//		String[] lines = contents.split("\n");
		List<String> texts = makeMultiContents(contents);
//		for (String str : texts) {
//			System.out.println("=====" + str);
//		}

		for (int x = 0; x < texts.size(); x++) {
			mTargetSheet.insertRow(index + 1 + x);// 新添加一行，需要加在原来行的下面
			for (int i = 0; i < 21; i++) {// 复制上面一行
				Cell cell2 = mTargetSheet.getCell(i, index);
				contents = cell2.getContents();
				// System.out.println("======\n" + contents);
				CellFormat cellFormat = cell2.getCellFormat();
				String txt = contents;
				if (i == 7) {
					txt = texts.get(x);
				}
				Label l = new Label(i, index + 1 + x, txt);
				l.setCellFormat(cellFormat);
				mTargetSheet.addCell(l);
			}
		}

	}

	private void runTask() throws Exception {
		mTargetSheet = getTargetSheet();
		int rows = mTargetSheet.getRows();
		for (int i = rows - 1; i > 0; i--) {
			addAtRowIndex(i);
			//需要删除原来的数据，否则出现一条重复的数据
			mTargetSheet.removeRow(i);
		}
		System.out.println("=====over");
//		for (int i = 1; i < 10; i++) {
//			Cell cell = mTargetSheet.getCell(7, i);
//			String contents = cell.getContents();
//			System.out.println("======\n" + contents);
////			CellFormat cellFormat = cell.getCellFormat();
////			Label l= new Label(0,0,"aaaaa");
////			l.setCellFormat(cellFormat);
////			mTargetSheet.addCell(l);
//		}
		mWorkbook.write();
		mWorkbook.close();
	}

	private WritableSheet getTargetSheet() throws Exception {
		File xmlFile = new File("z.xls");
		Workbook rwb = Workbook.getWorkbook(xmlFile);
		mWorkbook = Workbook.createWorkbook(new File("target.xlsx"), rwb);
		return mWorkbook.getSheet(0);
	}

	private void test() throws Exception {
		File xmlFile = new File("a.xlsx");
		Workbook rwb = Workbook.getWorkbook(xmlFile);
		WritableWorkbook workbook = Workbook.createWorkbook(new File("b.xlsx"), rwb);
		WritableSheet sheet = workbook.getSheet(0);
		Cell cell = sheet.getCell(0, 0);
		Label l = new Label(0, 0, "aaaaa");
		sheet.addCell(l);
		workbook.write();
		workbook.close();
	}

	private void open() throws Exception {
		System.out.println("open====");
		File xmlFile = new File("a.xlsx");
		WritableWorkbook workbook = Workbook.createWorkbook(xmlFile);
		WritableSheet sheet = workbook.createSheet("sheee1", 0);
		for (int row = 0; row < 10; row++) {
			for (int col = 0; col < 10; col++) {
				sheet.addCell(new Label(col, row, "data" + row + col));
			}
		}

		WritableSheet sheet1 = workbook.createSheet("sheee12222", 0);
		for (int row = 0; row < 10; row++) {
			for (int col = 0; col < 10; col++) {
				sheet1.addCell(new Label(col, row, "data" + row + col));
			}
		}
		workbook.write();
		workbook.close();
	}

	String shu[] = { "语音录入", "选择结果" };
	private List<String> result = new ArrayList<String>();
	public Stack<String> stack = new Stack<String>();

	private void f(String[] shu, int targ, int cur) {
		if (cur == targ) {
			// System.out.println(stack);
			StringBuilder sb = new StringBuilder();
			for (String s : stack) {
				sb.append(s);
				sb.append(",");
				// System.out.println("===x=" + s);
			}
			String str = sb.toString();
			result.add(str.substring(0, str.length() - 1));
			return;
		}
		for (int i = 0; i < shu.length; i++) {
			stack.add(shu[i]);
			f(shu, targ, cur + 1);
			stack.pop();
		}
	}

}
