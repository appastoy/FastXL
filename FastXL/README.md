# FastXL

This project is a fast excel file reader C# library. This supports only read. NOT write.

[![.NET Test](https://github.com/appastoy/FastXL/actions/workflows/dotnet_test.yml/badge.svg?branch=develop)](https://github.com/appastoy/FastXL/actions/workflows/dotnet_test.yml)
[![.NET Publish](https://github.com/appastoy/FastXL/actions/workflows/dotnet_publish.yml/badge.svg?branch=master)](https://github.com/appastoy/FastXL/actions/workflows/dotnet_publish.yml)
[![NuGet version (AppAsToy.FastXL)](https://img.shields.io/nuget/v/AppAsToy.FastXL.svg?style=flat-square)](https://www.nuget.org/packages/AppAsToy.FastXL/)

__GitHub__: https://github.com/appastoy/FastXL

__NuGet__: https://www.nuget.org/packages/AppAsToy.FastXL


# How to
Excel file(.xlsx, .xlsm) is a set of XML files that are compressed to zip archive.
This library reads XML files in zip archive directly. 
Especially, For performance, It uses XmlReader instance to parsing forward. 
So you can access minimum things of excel file specification.
