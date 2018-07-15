package com.layton.matthew;

import org.kohsuke.args4j.Option;
import org.kohsuke.args4j.spi.StringArrayOptionHandler;

import java.util.List;

public class TimesheetReaderOptions {
    protected String readLocation = "TimesheetTemplate.docx";
    @Option(name="--write",  usage="Where to write output")
    protected String writeLocation;
    @Option(name="--start", handler = StringArrayOptionHandler.class, usage="Comma delimited list of start times")
    protected List<String> startTimes;
    @Option(name="--end", handler = StringArrayOptionHandler.class, usage="Comma delimited list of end times")
    protected List<String> endTimes;
    @Option(name="--lunch",handler = StringArrayOptionHandler.class, usage="Comma delimited list of lunch times")
    protected List<String> lunchTimes;
}
