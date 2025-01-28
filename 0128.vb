private void button1_Click(object sender, EventArgs e)
{
    // 获取下拉框选中项的文本
    string selectedItem = comboBox1.SelectedItem?.ToString();

    // 判断是否有选中项
    if (selectedItem != null)
    {
        MessageBox.Show($"你选择了：{selectedItem}");
    }
    else
    {
        MessageBox.Show("请先选择一个选项！");
    }
}
