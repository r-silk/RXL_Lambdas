# RXL_Lambdas
A repository for developing, storing, and sharing Excel LAMBDA functions

In this repository the custom Lambda functions are stored within Excel files (.xlsx). The files include some documentation anmd examples of the development steps, a final version of the full function along with full function documentation (e.g. purpose, parameters, error codes), and some examples of the intended function usage.

If you wish to use an RXL custom Lambda function within your own Excel Workbook, there are 3 main options available to you:\
1. Each Excel file contains a copy of the final function with it's appropriate name within the Name Manager which can be copied across to a new Workbook as you normally would for copying a defined name between Workbooks.
2. Each Excel file also contains a "minimised" version of the final funciton - this is copy of the final function (as displayed in the file) but with all unnecessary spaces and any code comments removed and as a single text string. This minimised version can be copied and pasted directly into the Name Manager when creating a new defined name via the "Refers to:" input box in the "New Name" window.
> [!NOTE]
> When adding the functions in this manner please use the suggested function name all of which will start with `RXL_`.\
> Please note that the non-minimised version of the final function cannot be directly copied to the Name Manager as displayed as it contains code-comments and the spacing/indentation used may cuase the function text string to exceed the 2088 character limit in the Name Manager's "Refers to:" input box.
3. Finally, at the top of the 'Demo' worksheet there will be a link to a GitHub Gist where a copy of the function can be found in a code-commented format along with full function documentation. These versions of the RXL functions can be uploaded and added to your Workbook using the Excel Labs Advanced Formula Editor (AFE) using the Gist URL. The AFE versions of the function are added and stored in your Workbook with the code-commenting and documentation intact, and the correct suggested name, so they can be referenced easily via the AFE at any time (including when offline).
> [!TIP]
> There are also RXL function modules available as Gist files which can be added to your Workbook with multiple similar funcitons stored together. These modules are only available as Gists; if using one of the other above methods you will need to add each of the module functions individually. 



R Silk\
RXL Development, Chelmsford, UK\
https://excel-bits.net
