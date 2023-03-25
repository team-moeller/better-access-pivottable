# Better-Access-PivotTable
Better pivot tables for Access with pivottable.js

## Why Better PivotTable for Access?

Microsoft Access has lost the ability to create pivot tables and pivot charts with version Access 2013. The official recommendation is to use Excel for this purpose. 

But why use Excel when we can have it in Access. There are many solutions for pivot tables based on Java Script available on the web. This project makes use of this.
We create pivot tables using the Pivottable.js library and display them in the web browser control. The whole logic is hidden in a class module.

Take a look at the demo and let yourself be inspired by the possibilities.

## You want to give it a try?
1. Download the [latest release](https://github.com/team-moeller/better-access-pivottable/releases/latest)
2. Unpack the files to a trusted folder
3. Run the database
4. Push the button: "Create Pivot table"

## How to integrate into your own database?
**1. Import of the class modules**

First, all modules with the name "BAPT_*" must be imported from the demo database into your Access database.

**2. Insert web browser control on form**

The second step is to add a web browser control to display the chart on a form. It is best to give the control a meaningful name. This is required later in the VBA code. I like to use the name "ctlWebbrowser" for this.

The following text is entered in the "ControlSource" property: = "about: blank". This ensures that the web browser control remains empty at the beginning.

**3. First lines of code for the basic functionality**

The best thing to do is to add another button. In the click event, paste the following code:

```vba
Dim myPivot As BAPT_PivotTable  
Set myPivot = BAPT.PivotTable(Me.ctlWebbrowser)
myPivot.ShowPivot
```

* In line 1 a variable of the type BAPT_PivotTable is declared.
* In line 2 a new instance of this class is created and the web browser control is assigned to the class module.
* The pivot table is created in line 3. 

When you run this code, you will see a pivot table with some data. At the moment no data source is assigned. In such a case, Better-Access PivotTable simply shows a standard data source with 6 entries. This is particularly practical for our example. We have now done a quick test and fundamentally implemented the pivot table.

**4. Add a data source and define some basic attributes**

In order for the pivot table to show something, it needs a [data source](https://github.com/team-moeller/better-access-pivottable/wiki/datasource). You can use the [DataSource.ObjectName](https://github.com/team-moeller/better-access-pivottable/wiki/datasource#objectname) property for this, for example. Enter the name of a table or a query that contains the data to be displayed.

Finally, use the [Aggregator](https://github.com/team-moeller/better-access-pivottable/wiki/pivottable#aggregator) and the [AggregatorField](https://github.com/team-moeller/better-access-pivottable/wiki/pivottable#aggregatorfield) properties specify how the data is initially summarized in the pivot table. Use the [RendererName ](https://github.com/team-moeller/better-access-pivottable/wiki/pivottable#renderername) property to define how the data is initially displayed.

The necessary VBA code could look like this, for example:

```vba
myPivot.DataSource.ObjectName = "tbl_DemoData"
myPivot.Aggregator = "Sum"
myPivot.AggregatorField = "ExtendedPrice"
myPivot.RendererName = "Table"
```

* In line 1, the table "tbl_Demo_Data" is specified as the data source.
* Line 2 sets the aggregator to "sum".
* Line 3 defines the aggregator field.
* In line 4, the renderer name is set.

**5. Set further attributes for the pivot table**

The next step is to adapt the pivot table to your own needs. For example, you can define the filed names for [rows](https://github.com/team-moeller/better-access-pivottable/wiki/rows) and [columns](https://github.com/team-moeller/better-access-pivottable/wiki/columns) or adjust the some [styles](https://github.com/team-moeller/better-access-pivottable/wiki/styles).

The project currently has 25 subclasses with lots of properties. You can see all of them in the documentation on the [Wiki](https://github.com/team-moeller/better-access-pivottable/wiki). I have also presented the individual progress in the [blog](https://blog.team-moeller.de/search/label/Better%20Access%20PivotTable).

As you can see, there are a multitude of sources. Take a look around and make use of the options provided.
