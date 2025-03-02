# How to Retrieve and Fill in Image Properties (Width, Height, Top, Left)

When using the Excel to PowerPoint Exporter tool, you'll need to specify the exact dimensions and positioning for each image being transferred to your PowerPoint presentation. This guide will show you how to easily retrieve the required image properties (width, height, top, left) using the tool and correctly enter them into the export table.

#### Step-by-Step Instructions

1.  **Open the PowerPoint Presentation**

    Start by opening the PowerPoint presentation that already contains the images you previously inserted using the Excel to PowerPoint tool.
2.  **Select the Image**

    Navigate to the slide containing the image you wish to adjust. Click on the image to select it. You should see the image highlighted with resize handles, indicating it is selected.
3.  **Click the 'Get Image Properties' Button**

    Now, go back to the Excel workbook that contains the Excel to PowerPoint tool. Click on the button labeled **"Get Image Properties"**. This button is designed to extract the width, height, top, and left position of the selected image in PowerPoint.
4.  **View and Copy the Image Properties**

    Once you click the "Get Image Properties" button, a message box will appear displaying the following details about the selected image:

    * **Width**
    * **Height**
    * **Top** (the distance from the top of the slide)
    * **Left** (the distance from the left side of the slide)

    More importantly, these values are automatically copied to your clipboard as an array for easy pasting.
5.  **Paste the Properties in Excel**

    Next, go back to the export table in the Excel workbook. Locate the row corresponding to the named range you are working with, and click on the **"Width"** cell for that named range.

    * When you paste the values into the "Width" cell, it will automatically paste the entire array, including the **Height**, **Top**, and **Left** values, into their respective cells in the same row.
6.  **Repeat for Additional Images**

    If you have multiple images, simply repeat the steps above for each image to ensure the positioning and size are exactly as you desire.

{% hint style="info" %}
Note: Make sure to **select the correct image in PowerPoint** before clicking the "Get Image Properties" button. If no image is selected, you will receive an error message.
{% endhint %}

See the demo animation below for a visual walkthrough of these steps, which makes the process even clearer and more straightforward.

{% embed url="https://iframe.mediadelivery.net/play/289332/ff931080-81e7-474c-8e7d-ec967ad6c837" %}
Get Image Properties Demo
{% endembed %}

