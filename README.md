# ADVL_Utilities_Library_1
Library of utility classes used by Andorville™ applications.


- - -
Copyright 2016 Signalworks Pty Ltd, ABN 26 066 681 598

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.



- - -

The Andorville™ Utilities Class Library contains a set of common classes used by Andorville™ applications.
It contains classes to manage projects, compress data, and run XMessage instructions used to communicate between applications.

**Classes currently include:**

**ZipComp** - used to compress files into and extract files from a zip file.

**XSequence** - runs XSequence files. (An XSequence is an AL-H7™ Information Vector Sequence stored in an XML format. An Information Vector contains an Information value and a Location value. An Information Vector Sequence can be used to store processing steps, data sets and mathematical expressions.)

**XMessage** - runs XMessage instructions. (An XMessage is a simplified XSequence used to exchange information between applications.)

**Project** - this class manages Andorville™ Projects. These are directories or archive files created by an Andorville™ Application to store data related to a project.

