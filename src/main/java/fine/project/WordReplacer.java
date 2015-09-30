package fine.project;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * Utility class to replace string in word document
 * 
 * @author Vroman
 *
 */
public class WordReplacer {
	private XWPFDocument document;
	private String filePath;

	public WordReplacer(String filePath) {
		this.filePath = filePath;
	}

	public void readDocument() throws IOException {
		document = new XWPFDocument(new FileInputStream(filePath));
	}

	public void replaceText(Map<String, String> vals) {

		for (XWPFParagraph paragraph : document.getParagraphs()) {
			for (XWPFRun runs : paragraph.getRuns()) {
				String text = runs.getText(0);
				String replacedText = text;
				String[] searchList = vals.keySet().toArray(new String[vals.size()]);
				String[] replacementList = vals.values().toArray(new String[vals.size()]);
				text = StringUtils.replaceEachRepeatedly(replacedText,searchList ,replacementList);
				runs.setText(text, 0);
			}
		}
	}

	public void saveWord(String filePath) throws FileNotFoundException, IOException {
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(filePath);
			document.write(out);
		} finally {
			out.close();
		}
	}
}