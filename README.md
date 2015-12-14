# MeadCo ScriptX .NET Console Application Samples
Sample console application(s) illustrating the use of licensed ScriptX to print html documents.

## v1.0 ##
A single application based on an application that we use in our post-build verification test suite. The application illustrates the use of an application license for ScriptX and the use of the PrintHTMLEx() api to download and print html documents from a web site with custom headers for each document printed.

The printing code is simple. Various utility classes are used to simplify instantiation of the ScriptX COM objects, apply the license and dispose of the objects in a timely manner.

The sample is provided as a Visual Studio 2013 Solution for .NET 4.5. The code can be built with earlier versions of Visual Studio and .NET (e.g. 3.5).

Instructions on building with earlier versions of .NET are provided in the source Program.cs. To use an earlier version of Visual Studio, create a new project and add the files to the project. 

### ScriptX License ###
Use is made of licensed functions provided by the ScriptX Add-on. The sample code shows the use of a test license. To run the compiled code for yourself, please contact us at feeback@meadroid.com to obtain an evaluation license.

### Source License ###
The source code is unencumbered by any license. Use any of the source code in any way you wish.
 
