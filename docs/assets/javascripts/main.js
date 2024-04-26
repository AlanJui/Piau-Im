$(document).ready(function () {
    // 初始設定：隱藏所有文章的注音方式列表
    $('#articles ul').hide();

    function bindEvents() {
        // 解綁已有的事件，避免重複綁定
        $('#articles').off('click', 'h2');
        $('#articles').off('click', 'ul a');

        // 綁定點擊事件到所有標題上，使用委派確保動態元素也能綁定事件
        $('#articles').on('click', 'h2', function () {
            var phoneticList = $(this).next('ul');
            phoneticList.toggle();
        });

        // 綁定點擊事件到文章連結上
        $('#articles').on('click', 'ul a', function (e) {
            e.preventDefault();  // 阻止默認行為
            var url = $(this).attr('href');  // 獲取連結的 URL
            $('#main').load(url);  // 載入新內容
        });
    }

    // 初始時綁定事件
    bindEvents();

    // 存儲原始的 articles 內容
    var originalArticles = $('#articles').clone(true);

    // 點擊 "回到首頁" 時的處理程序
    $('nav ul li:first-child a').on('click', function (e) {
        e.preventDefault();  // 阻止默認行為
        $('#main').empty();  // 清空主要內容區域
        $('#main').append(originalArticles);  // 添加儲存的原始 articles 內容
        $('#articles ul').hide();  // 重設文章列表為隱藏狀態
        $('nav ul').hide();  // 確保菜單被收合
        bindEvents();  // 重新綁定事件
    });

    // 點擊 menu 按鈕時的處理程序
    $('nav button').on('click', function () {
        $('nav ul').toggle();  // 切換菜單的顯示狀態
    });
});
