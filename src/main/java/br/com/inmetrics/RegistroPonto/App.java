package br.com.inmetrics.RegistroPonto;

import org.junit.internal.TextListener;
import org.junit.runner.JUnitCore;
import org.junit.runner.Result;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
    	  JUnitCore junit = new JUnitCore();
    	  junit.addListener(new TextListener(System.out));
    	  Result result = junit.run(CapturaPonto.class); // Replace "SampleTest" with the name of your class
    	  if (result.getFailureCount() > 0) {
    	    System.out.println("Test failed.");
    	    System.exit(1);
    	  } else {
    	    System.out.println("Test finished successfully.");
    	    System.exit(0);
    	  }
        System.out.println( "Hello World!" );
    }
}
