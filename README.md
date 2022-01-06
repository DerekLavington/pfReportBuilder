# pfReportBuilder
A report builder in MS Word VBA

So this project is about building an app to assist (as opposed to automating) the building of professional reports. While intended as a general tool, it's inspiration and design reflects regulatory requirements relating to investment advice that require financial advisers to issue a 'suitability report' to customers. Reference to this requirement and the wider context relating to suitability reports can be found at https://www.handbook.fca.org.uk/handbook/COBS/9/?view=chapter under section 9.4.

Given that preparing such reports can be complex and time consuming, there is a need for financial advisers to seek efficiencies. Typically, this tends either to be 'model reports' or utilisation of previous reports as models. However, this can be problematic insofar as 

1) Reports produced tend to be insufficiently personalised and tend therefore to be non-compliant with regulatory requirements.

2) Errors in the model flow through to new reports, particularly where based on previous client reports. 

3) The effort involved in improving, extending and updating models often results in 'lack of ownership' of content, creating stagnation in the model that progressively affects the quality of reports. 

5) Models tend to become lengthy and difficult to maintain when catering for the range of required scenarios, particularly in relation to complex financial planning requirements. 

These issues arise largely because most financial advisers use static as opposed to automated templating. While automation software is available, it tends to bundled with expensive back-office systems that may not be practical, or commercial for the adviser to adopt. Such systems also tend to focus on providing content wording as opposed to assisting the adviser in developing and using their own wording. This may be acceptable where the focus is on the sale of simple financial products, but is not sufficient to support practitioners who engage in complex and detailed financial planning for high net worth customers.

The objective of this project is, therefore, to provide an automation tool that enables financial adviser to develop and maintain their own content wording when constructing suitability reports. It does so by allowing optional content to be selected during report production while, at the same time, providing for the underlying template to be easily updated and extended to cater for new options. The project uses MS Word as a platform as this software is ubiquitous within most financial services organisations.  

Functionality of pfReportBuilder breaks down into modules that allow the user to assemble a series of pre formed report sections into a template, from which selection of personalised options can then be automated in order to produce a final report. Release 2.1 provides 'proof of concept' functionality for assembling the pre formed report sections into a template. There is also a branch from the main code development called ReportMangler, which develops the functionality by which options within the pre formed report sectons can the be personalised to automate the final report.  

the pfReportBuilder code base is now being developed towards version 3.0, which will combine the ReportMangler functionality and provide the first beta release. In addition, documentation providing necessary background to the theory and practice of suitability reports, plus use of pfReportBuilder software, will be included.
