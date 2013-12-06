# J

Simple data wrapper that attempts to wrap [xlsjs](http://npm.im/xlsjs) and [xlsx](http://npm.im/xlsx) to provide a uniform way to access data

The key function is `J.readFile`, which takes a filename and returns an array.  The first object in the array is the module corresponding to the file type, and the second object is the actual content.
