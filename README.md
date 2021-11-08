# excel-link-remover-utilities
A random assortment of utilities that might be helpful for broken links in Microsoft Excel (as of ~2021)

# Problem
![image](https://user-images.githubusercontent.com/75244862/140740409-76df2a20-a9b2-4201-af19-38c6dd6ca74d.png)
Sometimes, when opening a file, Excel complains that it cannot find links to files. But as far as you - the user - know, there are no links in formulas present and Excel may or may not even show a link in the link manager. The same problem arises that you **cannot break the link** with the normal button from the link manager, so you're stuck with the non-existent connection.
![image](https://user-images.githubusercontent.com/75244862/140740421-7ba11566-6516-432a-b610-1c22c83d4486.png)

# Normal Solution: Hidden Names

What often is the problem in this cases is that worksheets were **copied **over from other files. While copying those worksheets, the **named ranges** will stay intact as well (this happening with most of our Finance templates as we're often using named ranges). This is a problem for other users, because sometimes those named ranges will still refer back to the old file and, therefore, create an external link. These external links are not breakable as they are not in cells for Excel to access.

Sometimes these named ranges are even hidden, which is why we need an AddIn to deal with this problem. Download the file **Name Manager 2007.xlam** and copy it into the following folder:

`%appdata%\Roaming\Microsoft\AddIns`

After you copied the file, you also need to enable the add-in in Excel settings. For this go to **Files --> Options --> AddIns --> Go... --> Enable Name Manager Utility**
![image](https://user-images.githubusercontent.com/75244862/140740438-a1641398-29b0-4298-8992-60eaa367d643.png)

Now, you go to the **Formulas** tab and click on **Name Manager:**
![image](https://user-images.githubusercontent.com/75244862/140740457-076000db-d41c-42c2-b5e5-bca1439ac843.png)

In the following opening window, you're now able to see all named ranges - the ones visible and invisible to you. To filter for all ranges that are causing problems, you can select **Or** on the right and then select **external references** and **errors** (select both with Ctrl pressed). Now you only see named ranges that are causing problems:
![image](https://user-images.githubusercontent.com/75244862/140740465-ae80be40-d6ee-4ee4-9491-8488331274f3.png)

Now simply select all named ranges that are visible to you by clicking the first, holding **Shift** and pressing **End** on your keyboard or hold **Shift**, then scroll down and **click** on the last entry.
![image](https://user-images.githubusercontent.com/75244862/140740475-d8746874-dc43-4c25-b0fc-be48fb6fc489.png)

In the upcoming dialog fields, select **Yes** (you want to really delete those names) and **No** (you do not want to replace them with references). This may take a while and seem like it hangs while it's processing but should finish after a few minutes.

Now all references that are hidden in names are deleted. In 90% of all cases, this should result in the link also being deleted. If it is still there in the link manager, try to save the file (**as a new version!**) and open it up again. If the link then is still present, we need to move to the advanced solution below.

# Advanced Solution: Links in Data Validation Settings

This problem has occured in the Analytical Package template. Here, an external link was hidden in the data validation settings of some cells, therefore the link still was not breakable but was also not removed with the Name Manager AddIn. Here we need the big guns, [[attach:findlink.xlam||target="_blank"]]. Also download this and save it in the AddIns folder:

`%appdata%\Roaming\Microsoft\AddIns`

After you copied the file, you also need to enable the add-in in Excel settings. For this go to **Files --> Options --> AddIns --> Go... --> Enable Findlink**
![image](https://user-images.githubusercontent.com/75244862/140740491-0f3e1269-74b8-4293-bce0-2baa5218b8ed.png)

Now, we can go to the newly created button **Add-Ins --> Find Links**. Here you are now able to select the link that is causing problems and select what you want to do. You have three choices: Ask for each separate link (a bit of a hassle), search for the links only (to see where the problem comes from) or - what we need - search for the links and delete the links, so we select the third option and run it:
![image](https://user-images.githubusercontent.com/75244862/140740504-6951499b-0f78-4526-bad1-e49b3ffe9bb8.png)

This also takes a while and you might have to wait a little as it goes through each sheet and tries to find the culprit links. But after it is done, it opens a new Excel workbook and gives you a report over what it has done:
![image](https://user-images.githubusercontent.com/75244862/140740518-09668892-88a0-4e42-9d7d-8062caf2ffb3.png)

Now, you can save the file again (as a new version, please, as you don't know what exactly it has done) and after opening it again, the last link should have been removed from the link manager. Check in **Data --> Edit Links** if that is the case. Also check your file - especially the data validation fields - if you have created a mess of formats or some data looks weird now.
