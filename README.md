# PowerPointGenerator

This tool provides an easy way to generate a powerpoint with slides based on a CSV file.
The pptx files used as a template has to provide one slide with at least one placeholder.
The placeholders are formated like this: ##FIELDNAME##
The Search-Alorythm is Regex-based an searches a Shape with includes a placeholder. All text within the shape is replaced.
The CSV File has to provide a header row. All Headers are handled as fields. Whitespaces are ignored.
Headers should be distinct. 