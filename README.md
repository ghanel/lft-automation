# lft-automation

/**
 * Lateral flow testing AutoIT script generator
 * Generates AutoIT scripts for lateral flow tests.
 * Can be used for UK Government Website
 *
 * This script can be run on any computer however the resulting script will have to be run on a Windows computer with AutoIT installed. 
 * https://www.autoitscript.com/site/
 * 
 * The Windows computer should also have Google Chrome browser installed 
 * 
 * Items that you may wish to change are indicated by comments
 * 
 * 
 * 
 * 
 * 
 * ENSURE YOU CHANGE YOUR SITE ID!!!  Line 119
 * 
 * 
 * Current limitations (because we haven't found them necessary or government didn't ask us to collect the information)
 * Futher info on Ethnicity not input (only Asian/Black/White/Mixed/Other. No follow up for White of British/Irish etc). Left as "Prefer not to say"
 * "Prefer not to say" entered for travel to workplace question
 * Country is hard coded to England change 
 * Change line 270 to
 * Send("{TAB}{TAB}{TAB}{TAB}{DOWN}{TAB}{ENTER}") for Scotland
 * Send("{TAB}{TAB}{TAB}{TAB}{DOWN}{DOWN}{TAB}{ENTER}") for Northern Ireland
 * Send("{TAB}{TAB}{TAB}{TAB}{DOWN}{DOWN}{DOWN}{TAB}{ENTER}") for Wales
 * 
 * Email address and mobile number are required
 * Landline number is not collected or input
 * NHS number is not collected or input
 * Vaccine question is hard coded to "No"
 * 
 * 
 * If you find this useful, please visit www.id3.org.uk and follow @id3company 
 * or follow me @gavinhanel
 *
 * @author Gavin Hanel https://www.id3.org.uk gavin@id3.org.uk  
 * @lastmodified 25 January 2021
 * @version 1 Gavin Hanel
 */

/* @license MIT License

Copyright (c) 2021 H&H Technology Ltd

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/
