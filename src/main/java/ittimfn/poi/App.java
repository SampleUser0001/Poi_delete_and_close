package ittimfn.poi;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

/**
 * Hello world!
 *
 */
public class App {

    private static Logger logger = LogManager.getLogger(App.class);

    public static void main( String[] args ) throws Exception {
        logger.trace("Delete column args.length : {}", args.length);
        logger.info("Delete column args[0] : {}", args[0]);
        logger.info("Delete column args[1] : {}", args[1]);
        new DeleteColumn().delete(args[0], args[1]);
    }
}
