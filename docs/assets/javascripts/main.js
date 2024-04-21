// JavaScript 代碼將在這裡添加
$(document).ready(function () {
    // 隱藏所有文章的注音方式列表
    $('#articles ul').hide();

    // 為每個標題添加 click 事件
    $('#articles h2').click(function () {
        // 找到被點擊的標題下的注音方式列表
        var phoneticList = $(this).next('ul');

        // 切換注音方式列表的顯示狀態
        phoneticList.toggle();
    });
});
