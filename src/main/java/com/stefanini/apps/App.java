package com.stefanini.apps;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) throws Exception
    {
    	FeedbackJoiner fj = new FeedbackJoiner();
    	fj.join(args.length>=1?args[0]:null,
    			args.length>=2?args[1]:null);
    }
    
    
}
