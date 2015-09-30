package fine.project;

import java.util.Map;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Options;

/**
 * Main class
 * 
 * @author Vroman
 *
 */
public class App {

	private static Options options = new Options();
	static {
		options.addOption("tw", true, "Word template location");
		options.addOption("ex", true, "Excel location");
		options.addOption("rw", true, "Result Word location");
		options.addOption("h", false, "help");
	}
	private static HelpFormatter helpFormatter = new HelpFormatter();
	private static CommandLineParser parser = new DefaultParser();
	
	public static void main(String[] args) throws Exception {
		CommandLine cmd = parser.parse(options, args);
		
		if(args.length==0 || cmd.hasOption("h")){
			helpFormatter.printHelp("Fine Generator", options);
			return;
		}
		String templateLocation = cmd.getOptionValue("tw");
		String excellLocation = cmd.getOptionValue("ex");
		String resultLocation = cmd.getOptionValue("rw");
		
		
		Map<String, String> readKeys = XmlReader.readKeys(excellLocation);
		WordReplacer wordReplacer = new WordReplacer(templateLocation);
		wordReplacer.readDocument();
		wordReplacer.replaceText(readKeys);
		wordReplacer.saveWord(resultLocation);
	}

}
