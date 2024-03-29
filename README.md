# ucJLDatePicker
UserControl DatePicker GDI/GDI+ VB6
ucJLDTPicker is a user control created to be used in modern forms, which can be adjusted to some cooler layouts with its multiple properties.

This is its first version, which enters a trial period since it may present some undetected errors in the process, for this reason it will be appreciated if they can be reported.

This has a very important property, which allows us to use the control as a child of some form or float and linked to another control, this is called IsChild.

In the project, a control called ucText created by Leandro Ascierto has been used, which can be found on his website http://leandroascierto.com/blog/uctext-custom-texbox-unicode/


# Updates

**02/06/2022:**
  - The MaxRangeDays property has been enabled, with which we can limit the range of days to be selected.

**04/11/2022:**
  - Fixed some bugs in some properties.
  - New properties added.
    - **ShowTodayButton:** to enable the current day button when AutoApply is True.
    - **DayHotColor:** To control the color that will be displayed when the mouse is over someday.

**08/11/2022:**
  - Updated a bug in the use of FirstDayOfWeek.

**10/11/2022:**

  **First part**
  - Changed properties:
    - DayFreeForeColor To **DayHolidayForeColor**

  - Added properties:
    - ✔️ **DayHolidayBackColor**
    - ✔️ **DayReservedBackColor**

  - Removed properties:
    - ❌ DayOverBackColor
    - ❌ DayOverForeColor
  
  - New procedure:
    - **SetDaysRaservedAndHoliday:** Allows you to define the reserved days and holidays for subsequent counting among the selected ranges, to do this use the **CountFreeDay** or **CountReservedDay** properties, you can obtain the number of days selected with the **DaySelCount** output property.

  - New behavior:
    - Added the possibility of expanding an already selected range, you just have to right click on the end date and you can expand to the new date you need.

  **Second part**
  - **Errors found:**
    - Fixed bug in action button displays.
  
  - Changed name procedure:
    - SetRangeButtonsCaption To **SetRangeButtons**: The change is because the possibility of making the buttons visible or not was added, for this reason an optional **VISIBLE** parameter was added, which by default has the value **TRUE**
  
  - New behavior:
    - Range buttons will now close the control when their property **IsChild = False** and **AutoApply = True**
