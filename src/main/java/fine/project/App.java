package fine.project;

import java.util.Map;
/**
 * Main class
 * @author Vroman
 *
 */
public class App {
	public static void main(String[] args) throws Exception {
		
		Map<String, String> readKeys = XmlReader.readKeys("D:/lab2.xlsx");
		WordReplacer wordReplacer = new WordReplacer("D:/w.docx");
		wordReplacer.readDocument();
		wordReplacer.replaceText(readKeys);
		wordReplacer.saveWord("D:/wq.docx");
		
	}
}
