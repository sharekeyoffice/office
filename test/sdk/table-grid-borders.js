// Regression tests run with the pinned SDK's real document, styles and history.
$(function ()
{
	const logicDocument = AscTest.CreateLogicDocument();
	const sides = ["Top", "Bottom", "Left", "Right", "InsideH", "InsideV"];

	function createTable(styleId)
	{
		logicDocument.StartAction(AscDFH.historydescription_Document_AddTable);
		let table = logicDocument.AddInlineTable(3, 2);
		if (styleId)
			table.Set_TableStyle(styleId);
		logicDocument.FinalizeAction();
		return table;
	}
	function assertGrid(assert, table)
	{
		const expected = table.Get_Styles().Get_Pr(table.Get_TableStyle(), styletype_Table).TablePr.TableBorders;
		for (const side of sides)
		{
			assert.deepEqual(table.Pr.TableBorders[side], expected[side], side + " is explicit and matches the style");
			assert.notStrictEqual(table.Pr.TableBorders[side], expected[side], side + " does not share the style's border object");
		}
	}

	QUnit.module("Table Grid border compatibility", {
		beforeEach : function()
		{
			AscTest.ClearDocument();
			logicDocument.PushToContent(AscTest.CreateParagraph());
			logicDocument.MoveCursorToStartPos();
			logicDocument.History.Clear();
		}
	});

	QUnit.test("New Table Grid tables have explicit borders", function(assert)
	{
		const table = createTable();
		assert.strictEqual(table.Get_TableStyle(), logicDocument.Styles.Get_Default_TableGrid());
		assertGrid(assert, table);
	});

	QUnit.test("Nested tables have independent explicit borders", function(assert)
	{
		const outer = createTable();
		const inner = createTable();
		assert.notStrictEqual(inner, outer);
		assertGrid(assert, outer);
		assertGrid(assert, inner);
	});

	QUnit.test("Insertion and its borders undo and redo as one action", function(assert)
	{
		const table = createTable();
		logicDocument.Document_Undo();
		assert.strictEqual(logicDocument.GetElementsCount(), 1, "One undo removes the table");
		assert.ok(logicDocument.GetElement(0).IsParagraph());
		logicDocument.Document_Redo();
		assert.strictEqual(logicDocument.GetElement(0), table);
		assertGrid(assert, table);
	});

	QUnit.test("Reapplying Table Grid restores explicit borders and supports undo", function(assert)
	{
		const table = createTable();
		const none = new CDocumentBorder();
		none.Value = border_None;
		table.Set_TableBorder_Top(none);
		logicDocument.StartAction(AscDFH.historydescription_Document_ApplyTablePr);
		table.Set_TableStyle(table.Get_TableStyle());
		logicDocument.FinalizeAction();
		assertGrid(assert, table);
		logicDocument.Document_Undo();
		assert.strictEqual(table.Pr.TableBorders.Top.Value, border_None, "Undo restores the prior override");
		logicDocument.Document_Redo();
		assertGrid(assert, table);
	});

	QUnit.test("Preserving direct formatting keeps explicitly disabled borders", function(assert)
	{
		const table = createTable();
		const none = new CDocumentBorder();
		none.Value = border_None;
		table.Set_TableBorder_Top(none);
		table.Set_TableStyle(table.Get_TableStyle(), true);
		assert.strictEqual(table.Pr.TableBorders.Top.Value, border_None);
		assert.strictEqual(table.Pr.TableBorders.Bottom.Value, border_Single);
	});

	QUnit.test("Choosing a borderless style removes the materialized grid", function(assert)
	{
		const table = createTable(logicDocument.Styles.Get_Default_Table());
		for (const side of sides)
			assert.strictEqual(table.Pr.TableBorders[side], undefined, side + " is not forced");
		table.Set_TableStyle(logicDocument.Styles.Get_Default_TableGrid());
		assertGrid(assert, table);
	});

	QUnit.test("Unrelated table styles retain their own formatting", function(assert)
	{
		const style = new AscWord.CStyle("Custom borders", null, null, styletype_Table);
		style.Create_TableGrid();
		logicDocument.Styles.Add(style);
		const table = createTable(style.GetId());
		for (const side of sides)
			assert.strictEqual(table.Pr.TableBorders[side], undefined, side + " remains inherited");
	});

	QUnit.test("Loading existing Table Grid tables does not add direct borders", function(assert)
	{
		for (const flag of ["m_bLoad", "m_bRead"])
		{
			const table = AscTest.CreateTable(1, 1, [30]);
			logicDocument.PushToContent(table);
			const old = AscCommon.g_oIdCounter[flag];
			try
			{
				AscCommon.g_oIdCounter[flag] = true;
				table.Set_TableStyle(logicDocument.Styles.Get_Default_TableGrid());
			}
			finally
			{
				AscCommon.g_oIdCounter[flag] = old;
			}
			for (const side of sides)
				assert.strictEqual(table.Pr.TableBorders[side], undefined, flag + ": " + side);
		}
	});
});
