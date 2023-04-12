

let lastSelectedPath = "";

function 批量插入图_选中法() {
    let result = MsgBox("请先选中要插入图片的单元，已选中请点击【是】，未选中点击【否】", jsYesNo, "图片插入前提示");
    if (result !== 6) return;
    let pic_pick = Application.FileDialog(msoFileDialogFilePicker);
    pic_pick.AllowMultiSelect = true;
    pic_pick.InitialFileName = lastSelectedPath;
    if (pic_pick.Show() !== -1) return;
    lastSelectedPath = pic_pick.SelectedItems(1);
    let pic_geted = pic_pick.SelectedItems;
    let currentCell = ActiveCell;
    let direction = MsgBox("请选择插入方向，纵向请点击【是】，横向请点击【否】", jsYesNo, "图片插入方向");
    for (let counter = 1; counter <= pic_geted.Count; counter++) {
        let y = currentCell.Top;
        let x = currentCell.Left;
        let h = currentCell.Height;
        let w = currentCell.Width;
        let pic_Name = pic_geted.Item(counter);
        let obj = ActiveSheet.Shapes.AddPicture(pic_Name, msoFalse, msoTrue, x, y, w, h);
        obj.Placement = xlMoveAndSize;
        obj.LockAspectRatio = false;
        if (direction === 6) {
            currentCell = Cells(currentCell.Row + 1, currentCell.Column);
        } else {
            currentCell = Cells(currentCell.Row, currentCell.Column + 1);
        }
    }
}
