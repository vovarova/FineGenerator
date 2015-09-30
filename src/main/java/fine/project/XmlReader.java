package fine.project;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.function.Function;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.google.common.collect.ImmutableMap;

/**
 * Utility class that reads defined values from xcell
 * @author Vroman
 *
 */

public abstract class XmlReader {
	private static final Pattern LEFT = Pattern.compile("^l(#.*)$");
	private static final Pattern RIGHT = Pattern.compile("^r(#.*)$");
	private static final Pattern UP = Pattern.compile("^u(#.*)$");
	private static final Pattern DOWN = Pattern.compile("^d(#.*)$");
	private static final DataFormatter FORMATTER = new DataFormatter();
	private static final Map<Pattern, Function<Cell, String>> KEY_HANDLERS = ImmutableMap
			.<Pattern, Function<Cell, String>> builder()
			.put(LEFT, c -> FORMATTER.formatCellValue(c.getRow().getCell(c.getColumnIndex() - 1)))
			.put(RIGHT, c -> FORMATTER.formatCellValue(c.getRow().getCell(c.getColumnIndex() + 1)))
			.put(UP, c -> FORMATTER.formatCellValue(c.getRow().getSheet().getRow(c.getRowIndex() -1).getCell(c.getColumnIndex())))
			.put(DOWN, c -> FORMATTER.formatCellValue(c.getRow().getSheet().getRow(c.getRowIndex() + 1).getCell(c.getColumnIndex())))
			.build();

	public static Map<String, String> readKeys(String file)
			throws EncryptedDocumentException, InvalidFormatException, IOException {
		Map<String, String> keys = new HashMap<>();

		Workbook workbook = WorkbookFactory.create(new File(file));
		
		workbook.forEach(sh -> {
			sh.forEach(r -> {
				r.forEach(c -> {
					KEY_HANDLERS.forEach((p, f) -> {
						Matcher matcher = p.matcher(FORMATTER.formatCellValue(c));
						if (matcher.matches()) {
							keys.put(matcher.group(1), f.apply(c));
						}
					});
				});
			});
		});
		return keys;
	}
}
