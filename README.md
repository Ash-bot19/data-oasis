# ***📊 Restaurant Analytics Dashboard (Excel + VBA)***

I created multiple PivotTables to break down the dataset and replicate Zomato-style restaurant insights:

**1.Delivery Ratings Table (Green)**
- Compares average delivery ratings across restaurants in a selected city.
- Example: Aroma’s Hyderabad House in Pune consistently scored 4.2 on delivery for multiple bestseller items like Mutton Biryani and Chicken 65.

**2.Average Restaurant Rating Table (Orange)**
- Highlights the overall dining performance of restaurants.
- Helps identify consistency: e.g., Rolls Mania maintains a stable rating near 4.0, while Sukanta performs slightly better at 4.15.

**3.Cuisine-City Insights**
- Filtered via slicers and macros, so users can ask:
- “Which city has the best biryani delivery?”
- “Where do Chinese restaurants score higher — Mumbai or Delhi?”
---

## **🚀 Features**
- ✅ Interactive Dashboard – filter by cuisine, city, or both
- ✅ Zomato-like Functionality in Excel
- Check which cuisines are most popular in each city
  - Compare delivery ratings across cities and restaurants
  - Find the best performing restaurants by cuisine & location
- ✅ Checkbox + Macro Automation – dynamically connect/disconnect slicers to pivot tables
- ✅ Business Insights generated from raw data

## **🛠️ How It Works**

**1.Dataset**
- Restaurant and delivery dataset sourced from Kaggle.
- Contains details such as Restaurant Name, City, Cuisine, Delivery Rating, etc.

**2.Excel Analysis**
- PivotTables summarize cuisine performance by city.
- Delivery ratings compared across restaurants.
- Conditional formatting highlights top performers.

**3.Interactive Controls**
- Checkboxes linked to cells (TRUE/FALSE).
- VBA macro connects slicers to PivotTables based on checkbox state.
- Users simulate Zomato filters inside Excel (e.g., show me the best Chinese restaurants in Delhi).

## **📊 Key Insights**
Some of the insights uncovered through this dashboard include:
- 🍕 Cuisine Trends – Italian and Chinese dominate across metros, while regional cuisines perform better in smaller cities.
- 🚚 Delivery Ratings – Certain cities consistently score higher in delivery speed & quality (e.g., Delhi vs. Bangalore).
- ⭐ Top Performers – Some restaurants excel both in dine-in and delivery ratings, making them local favorites.
- 📍 City Comparison – Helps identify which city offers the best experience for specific cuisines (e.g., best biryani in Hyderabad).

## **⚙️ VBA Automation**
```vba
Sub ToggleSlicerConnection(SlicerName As String, PivotName As String, CheckCell As String)

    If Sheet1.Range(CheckCell).Value = True Then
        ActiveWorkbook.SlicerCaches(SlicerName).PivotTables.AddPivotTable _
            ActiveSheet.PivotTables(PivotName)
    Else
        ActiveWorkbook.SlicerCaches(SlicerName).PivotTables.RemovePivotTable _
            ActiveSheet.PivotTables(PivotName)
    End If

End Sub
```
This macro is assigned to each checkbox, allowing slicers to turn ON/OFF dynamically.

## **📂 Example Use Cases**
- ✅ Identify the best cuisine in each city
- ✅ Compare delivery performance across restaurants
- ✅ Check top-rated restaurants for a given cuisine
- ✅ Replicate Zomato-style filtering inside Excel
