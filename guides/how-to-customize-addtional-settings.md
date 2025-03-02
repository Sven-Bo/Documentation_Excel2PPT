# How to Customize addtional Settings

To customize the behavior of the Excel2PPT utility, you need to modify the settings found in the **SETTINGS** module of the VBA editor. If you are unsure how to access these settings, please refer to the following guide: [how-to-access-additional-settings.md](how-to-access-additional-settings.md "mention")

Below is a detailed explanation of each available setting:

#### 1. PowerPoint Stay Open After Export (`KEEP_POWERPOINT_OPEN`)

* **Description**: Determines whether PowerPoint remains open after the export completes.
* **Value Options**:
  * `True` - PowerPoint will stay open.
  * `False` - PowerPoint will close automatically after export (default).
* **Usage Scenario**: Set to `True` if you want to make further edits to the PowerPoint presentation immediately after export.

#### 2. Auto-Save PowerPoint Presentation (`AUTO_SAVE_PRESENTATION`)

* **Description**: Controls whether the exported PowerPoint presentation is saved automatically.
* **Value Options**:
  * `True` - The PowerPoint file will be automatically saved (default).
  * `False` - You will need to manually save the PowerPoint presentation.
* **Usage Scenario**: Set to `False` if you want to make manual adjustments to the presentation before saving it.

#### 3. Delay Before Pasting (`WAIT_TIME_BEFORE_PASTE`)

* **Description**: Defines the time (in seconds) to wait before pasting elements into PowerPoint. This delay helps prevent errors or crashes on slower systems.
* **Value Options**: Any integer value greater than or equal to `1`. Recommended: at least `1` second (default is `1`).
* **Usage Scenario**: Increase this value if you encounter issues with pasted content not appearing correctly, especially on slower computers.

#### 4. Maintain Aspect Ratio (`MAINTAIN_ASPECT_RATIO`)

* **Description**: Determines if the aspect ratio of the pasted shapes should be maintained during resizing.
* **Value Options**:
  * `True` - Aspect ratio will be maintained (default).
  * `False` - Shapes will be resized without maintaining the original aspect ratio.
* **Usage Scenario**: Set to `True` to ensure charts and graphics do not become distorted during resizing.

#### 5. Z-Order of Pasted Shape (`SHAPE_ZORDER`)

* **Description**: Controls the order of pasted shapes relative to other content on the slide.
* **Value Options**:
  * `"msoBringToFront"`: Brings the shape to the front.
  * `"msoSendToBack"`: Sends the shape to the back (default).
  * `"msoBringForward"`: Brings the shape one level closer to the front.
  * `"msoSendBackward"`: Sends the shape one level towards the back.
* **Usage Scenario**: Use this setting if you need to control how pasted shapes are layered in your slides, for instance, ensuring certain elements are in the foreground.

#### 6. Excel Table Name (`TABLE_NAME`)

* **Description**: Specifies the name of the table in your Excel worksheet where all the export configurations (e.g., ranges to copy, slide numbers) are listed.
* **Value Options**:
  * `"tbl_Excel2PPT"` (default) or any valid Excel table name.
* **Usage Scenario**: If you rename the table that holds your export configurations, update this setting to reflect the new table name.
