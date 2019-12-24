package com.asayclips;

import junit.framework.Test;
import junit.framework.TestCase;
import junit.framework.TestSuite;

/**
 * Unit test for simple App.
 */
public class AppTest 
    extends TestCase
{
    /**
     * Create the test case
     *
     * @param testName name of the test case
     */
    public AppTest( String testName )
    {
        super( testName );
    }

    /**
     * @return the suite of tests being tested
     */
    public static Test suite()
    {
        return new TestSuite( AppTest.class );
    }

    /**
     * Rigourous Test :-)
     */
    public void testFindOffsetDate() throws Exception
    {
        assertEquals("01/07/2018",  App.findOffsetDate("01/01/2018", 6) );
        assertEquals("01/01/2018",  App.findOffsetDate("01/14/2018", -13) );
        // daylight savings
        assertEquals("03/10/2018",  App.findOffsetDate("03/12/2018", -2) );
        assertEquals("03/12/2018",  App.findOffsetDate("03/10/2018", 2) );

        // daylight savings ends
        assertEquals("11/04/2018",  App.findOffsetDate("11/17/2018", -13) );
        assertEquals("11/17/2018",  App.findOffsetDate("11/04/2018", 13) );
        assertEquals("11/10/2018",  App.findOffsetDate("11/04/2018", 6) );
        assertEquals("11/04/2018",  App.findOffsetDate("11/10/2018", -6) );
    }


}
