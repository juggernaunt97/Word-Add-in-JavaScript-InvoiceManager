# [ARCHIVED] Word add-in: Loading data into custom XML parts bound to content controls in a Word document

**Note:** This repo is archived and no longer actively maintained. Security vulnerabilities may exist in the project, or its dependencies. If you plan to reuse or run any code from this repo, be sure to perform appropriate security checks on the code or dependencies first. Do not use this project as the starting point of a production Office Add-in. Always start your production code by using the Office/SharePoint development workload in Visual Studio, or the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), and follow security best practices as you develop the add-in. 


** Lưu ý: ** Kho lưu trữ này đã được lưu trữ và không còn được duy trì tích cực nữa. Các lỗ hổng bảo mật có thể tồn tại trong dự án hoặc các phần phụ thuộc của dự án. Nếu bạn định sử dụng lại hoặc chạy bất kỳ mã nào từ kho lưu trữ này, hãy đảm bảo thực hiện kiểm tra bảo mật thích hợp trên mã hoặc phần phụ thuộc trước tiên. Không sử dụng dự án này làm điểm bắt đầu của Phần bổ trợ Office sản xuất. Luôn bắt đầu mã sản xuất của bạn bằng cách sử dụng khối lượng công việc phát triển Office / SharePoint trong Visual Studio hoặc [Trình tạo Yeoman cho Phần bổ trợ Office] (https://github.com/OfficeDev/generator-office) và làm theo các phương pháp hay nhất về bảo mật khi bạn phát triển bổ trợ.

**Table of contents**

* [Summary](#summary)
* [Prerequisites](#prerequisites)
* [Key components of the sample](#components)
* [Description of the code](#codedescription)
* [Build and debug](#build)
* [Troubleshooting](#troubleshooting)
* [Questions and comments](#questions)
* [Contributing](#contribute)
* [Additional resources](#additional-resources)

<a name="summary"></a>
## Summary

In this sample we show you how to use the [JavaScript API for Office](https://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298(v=office.15)) to write data to a set of custom XML parts that are bound to content controls within a Word document. The following is a  picture of the scenario in question.

![Screenshot of running sample](https://cloud.githubusercontent.com/assets/8550529/9298298/4b980684-4461-11e5-8c00-8f86701e55c2.PNG)

We are creating packing slips from customer order data. The packing slip document is shown on the left of the preceding screen shot, with our Office Add-in on the right as a task pane app. When you select an order using the order id drop-down in the task pane on the right and then click the Populate button, the packing slip document is populated with data from that order.  The sample uses the Javascript API for Office to interact with the Word document by populating custom XML parts defined in the document with order data. These custom XML parts are bound to content controls that define the UI or the document. To simplify this sample, the order data is stored in the same JavaScript file that creates the add-in. However, in a real application, that data could come from a data source anywhere on the web.
Vietsub:
Chúng tôi đang tạo phiếu đóng gói từ dữ liệu đơn đặt hàng của khách hàng. Tài liệu phiếu đóng gói được hiển thị ở bên trái của ảnh chụp màn hình trước đó, với Phần bổ trợ Office của chúng tôi ở bên phải dưới dạng ứng dụng ngăn tác vụ. Khi bạn chọn đơn hàng bằng cách sử dụng id đơn đặt hàng -xuống dưới trong ngăn tác vụ ở bên phải và sau đó nhấp vào nút Điền, tài liệu phiếu đóng gói được điền dữ liệu từ đơn đặt hàng đó. Mẫu sử dụng API Javascript cho Office để tương tác với tài liệu Word bằng cách điền các phần XML tùy chỉnh được xác định trong tài liệu có dữ liệu đơn hàng. Các phần XML tùy chỉnh này liên kết với các điều khiển nội dung xác định giao diện người dùng hoặc tài liệu. Để đơn giản hóa mẫu này, dữ liệu đơn hàng được lưu trữ trong cùng một tệp JavaScript tạo bổ trợ. Tuy nhiên, trong một ứng dụng thực , dữ liệu đó có thể đến từ một nguồn dữ liệu ở bất kỳ đâu trên web.


<a name="prerequisites"></a>
## Prerequisites
This sample requires the following:  

  - Visual Studio 2013 with Update 5 or Visual Studio 2015.  
  - Word 2013 or later
  - Internet Explorer 9 or later, which must be installed but doesn't have to be the default browser. To support Office Add-ins, the Office client that acts as host uses browser components that are part of Internet Explorer 9 or later.
  - One of the following as the default browser: Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13, or a later version of one of these browsers.
  - Familiarity with JavaScript programming and web services.

<a name="components"></a>
## Key components

This solution was created in [Visual Studio](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS). It consists of two projects - InvoiceManager and InvoiceManagerWeb. Here's a list of the key files within those projects. 
#### InvoiceManager project

* [InvoiceManager.xml](https://github.com/OfficeDev/Word-Add-in-JavaScript-InvoiceManager/blob/master/InvoiceManagerSample/InvoiceManagerManifest/InvoiceManager.xml) The [manifest file](https://msdn.microsoft.com/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp) for the Word add-in.
* [PackingSlip.docx](https://github.com/OfficeDev/Word-Add-in-JavaScript-InvoiceManager/blob/master/InvoiceManagerSample/PackingSlip.docx) The example packing slip Word document used in this sample. 

#### InvoiceManagerWeb project

* [Home.html](https://github.com/OfficeDev/Word-Add-in-JavaScript-InvoiceManager/blob/master/InvoiceManagerSampleWeb/App/Home/Home.html) The HTML user interface for the Word add-in.
* [Home.js](https://github.com/OfficeDev/Word-Add-in-JavaScript-InvoiceManager/blob/master/InvoiceManagerSampleWeb/App/Home/Home.js) The JavaScript code used by Home.html to interact with Word using the JavaScript for Office API. 


<a name="codedescription"></a>
## Description of the code

For a detailed description of this sample, see [Exploring the JavaScript API for Office: Data Binding and Custom XML Parts](https://msdn.microsoft.com/en-us/magazine/dn166930.aspx)
## Mô tả mã

Để biết mô tả chi tiết về mẫu này, hãy xem [Khám phá API JavaScript cho Office: Liên kết dữ liệu và các phần XML tùy chỉnh] (https://msdn.microsoft.com/en-us/magazine/dn166930.aspx)

<a name="build"></a>
## Build and debug
1. Open the InvoiceManager.sln file in Visual Studio.
2. Press F5 to build and deploy the sample add-in and open it in Word.
3. On the **Home** ribbon, find the **Invoice Manager** group and press the **Open** button.
3. In the app task pane, select an order in the Order ID drop-down list.
4. Choose Populate to populate the packing slip in the Word document with information from the selected order.
<a name="build"> </a>
## Xây dựng và gỡ lỗi
1. Mở tệp InvoiceManager.sln trong Visual Studio.
2. Nhấn F5 để xây dựng và triển khai bổ trợ mẫu và mở nó trong Word.
3. Trên ruy-băng ** Trang chủ **, tìm nhóm ** Trình quản lý hóa đơn ** và nhấn nút ** Mở **.
3. Trong ngăn tác vụ ứng dụng, hãy chọn một đơn hàng trong danh sách ID đơn hàng thả xuống.
4. Chọn Populate để điền vào phiếu đóng gói trong tài liệu Word với thông tin từ thứ tự đã chọn.

You can view a list of the custom XML parts in a document by opening the XML Mapping pane in Word (Developer tab).

<a name="troubleshooting"></a>
## Troubleshooting

- If the add-in starts with a blank document, ensure that the **Start Document** property of the InvoiceManager project is set to *PackingSlip.docx* and not just to Word.
![before and after property settings page](https://cloud.githubusercontent.com/assets/8550529/9298211/b29908a8-445f-11e5-8887-0b3e6a9c8649.png)
- If the add-in does not appear in the task pane, Choose **Insert > My Add-ins >  InvoiceManagerSample**.

<a name="questions"></a>
## Questions and comments

- If you have any trouble running this sample, please [log an issue](https://github.com/OfficeDev/Word-Add-in-JavaScript-InvoiceManager/issues).
- Questions about Office Add-ins development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Make sure that your questions or comments are tagged with [office-addins].

<a name="contribute"></a>
## Contributing ##
We encourage you to contribute to our samples. For guidelines on how to proceed, see our [contribution guide](./Contributing.md)

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.


<a name="additional-resources"></a>
## Additional resources ##

- [More Add-in samples](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Office Add-ins](http://msdn.microsoft.com/library/office/jj220060.aspx)
- [Anatomy of an Add-in](https://msdn.microsoft.com/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp)
- [Bindings object (JavaScript API for Office)](http://msdn.microsoft.com/library/office/apps/fp160966.aspx)
- [Binding to regions in a document or spreadsheet](http://msdn.microsoft.com/library/office/apps/fp123511(v=office.15).aspx)
- [Creating an Office add-in with Visual Studio](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS)


## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.

# QTDA-CNTT nhóm A Plus
1. Nguyễn Minh Châu - 20158037
2. Lê Thị Liên - 20175978
3. Trần Thị Thủy-20176059
4. Lương Việt Anh - 20176004
5. Vũ Hữu Quốc Bảo - 20187160

