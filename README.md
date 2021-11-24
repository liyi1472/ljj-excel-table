# Excel 超级表切片器使用教程

业务场景要求维护一个多列数据表格。为了防止操作失误篡改原始数据，需要通过另一个独立的数据表格进行筛选。筛选时可以根据数据表格的列名指定条件，并可以级联过滤掉不存在的条件组，最终选出唯一的一条数据并可以在第三个独立的数据表格进行排版、展示和打印。

技术上要求使用 Excel 实现，不能使用 VBA（宏）编程，可适当使用 Excel 函数。

> 教程展示文件下载：[**demo.xlsx**](demo.xlsx)



## 所需知识

1. Excel 超级表的使用。
2. Excel 切片器的使用。
3. Excel 保护工作表的设置项。
4. Excel 相关函数的组合。



## 操作流程

1. 新建 Excel 工作簿，暂命名为 **demo.xlsx** 。

   <img src="images/01 新建Excel工作簿.png" width="30%"> 

2. **填充测试数据**。选中测试数据中的任意单元格，按 Ctrl+A 全选数据。

   <img src="images/02 填充测试数据.png" width="50%"> 

3. 按 **Ctrl+T** 将所选内容转换为超级表（也可以通过菜单“插入”→“表格”完成），弹出确认窗口点击“确定”。

   <img src="images/03 转换为超级表.png" width="50%"><img src="images/04 转换为超级表.png" width="50%"> 

4. 再次全选超级表中的内容，通过菜单**“插入”→“切片器”**，插入切片器。

   <img src="images/05 插入切片器.png" width="50%"> 

5. 新建数据表 Sheet2，将生成的切片器全部**剪切**到 Sheet2 按操作习惯排列整齐。

   <img src="images/06 移动切片器到Sheet2.png" width="80%"> 

6. 在 Sheet1 中最左侧插入新的一列，列名为**“序号”**，在A2单元格中输入以下公式，然后拖拽下拉应用到全部。

   ```
   =SUBTOTAL(3, Sheet1!$B$2:B2)
   ```

   应用该公式后，序号列会随筛选结果的变化自动变化显示序号。

   <img src="images/07 插入序号列.png" width="50%"> 

   同样是使用 **Ctrl+T** 将新列转换为超级表。

   <img src="images/08 序号列转换为超级表.png" width="50%"> 

7. 右键工作表 Sheet1，单击“保护工作表”，勾选**“使用自动筛选”**。

   这样就可以保证原始数据不会被误操作篡改，如需更新原始数据，右键工作表 Sheet1，单击“撤销工作表保护”即可。数据更新完成后，再次及时“保护工作表”即可。

   <img src="images/09 保护工作表.png" width="50%"> 

8. 在 Sheet2 中空白单元格内输入**“符合条件的结果数量：”**字样，并在紧邻右侧的单元格内输入以下公式。

   ```
   =SUBTOTAL(3, Sheet1!B:B)-1
   ```

   - [**SUBTOTAL()**](https://support.microsoft.com/zh-cn/office/subtotal-%E5%87%BD%E6%95%B0-7b027003-f060-4ade-9040-e478765b9939) 函数用来返回指定范围内的分类汇总数据。
   - [**COUNTA()**](https://support.microsoft.com/zh-cn/office/counta-%E5%87%BD%E6%95%B0-7dc98875-d5c1-46f1-9a82-53f3219e2509) 函数不会对空单元格进行计数，SUBTOTAL() 函数的第一个参数 "3" 代表使用 COUNTA() 统计指定范围内非空值单元格的数量，这样就能准确地获知根据切片器条件筛选的结果行数。
   - 再用非空单元格的数量减去表头单元格的数量1，当结果为1时，便可确保所筛选数据的唯一性。

   **点击 Sheet2 中切片器内的条件，即可完成对 Sheet1 中数据的筛选。**

   <img src="images/10 使用切片器筛选数据.png" width="80%"> 

9. 复制 Sheet1 中的表头部分粘贴到 Sheet2 中，建立筛选结果的**预览区域**。

   在表头每一列对应的列名下方单元格内输入以下公式。

   ```
   =VLOOKUP(1,Sheet1!$A:$E,2,FALSE)
   ```

   ```
   =VLOOKUP(1,Sheet1!$A:$E,3,FALSE)
   ```

   ```
   =VLOOKUP(1,Sheet1!$A:$E,4,FALSE)
   ```

   ```
   =VLOOKUP(1,Sheet1!$A:$E,5,FALSE)
   ```

   - [**VLOOKUP()**](https://support.microsoft.com/zh-cn/office/vlookup-%E5%87%BD%E6%95%B0-0bbc8083-26fe-4963-8ab8-93a18ad188a1) 函数用于在指定范围内按行查找首列中的指定内容，匹配后返回该行其他指定列的内容。

   - 在 Sheet1!\$A:\$E 范围内的A列序号列中查找序号为1的行，然后获取该行其他列的数据显示在预览区。
   - 对应的2列是“商场”，3列是“颜色”，4列是“种类”，5列是“价格”。

   <img src="images/11 显示筛选预览.png" width="80%"> 

10. 新建数据表 Sheet3，根据最终打印式样进行排版，将上一步中的对应公式填写到合适位置。

   <img src="images/12 打印排版.png" width="60%"> 

   <img src="images/13 打印预览.png" width="80%"> 

11. 为了在选择筛选条件的过程中能够**级联过滤掉不存在的条件组**，使用 Ctrl+A 选择全部切片器，选中后在切片器上右键，单击<span style="color:red;">**“切片器设置”**</span>，勾选<span style="color:red;">**“隐藏没有数据的项”**</span>。 

    <img src="images/14 切片器设置.png" width="80%"> 

    <img src="images/15 隐藏没有数据的项.png" width="80%"> 

    <img src="images/16 级联过滤条件组.png" width="80%"> 

12. 为了更加清晰地理解各表格的作用，分别右键 Sheet1、Sheet2、Sheet3 进行重命名。

    <img src="images/17 优化表名.png" width="40%"> 

