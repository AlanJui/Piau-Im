$(document).ready(function () {
    // 1. 文章列表搜尋功能
    $('#article-search').on('input', function() {
        const searchText = $(this).val().toLowerCase().trim();
        
        $('.card').each(function() {
            const articleTitle = $(this).find('.card-title').text().toLowerCase();
            
            // 修正：改用 class 來控制顯示/隱藏，以覆蓋 CSS 中的 !important
            if (articleTitle.includes(searchText)) {
                $(this).removeClass('hidden-card');
            } else {
                $(this).addClass('hidden-card');
            }
        });
    });

    // 2. 選單事件綁定
    function bindEvents() {
        $('#articles').off('click', 'h2');
        $('#articles').on('click', 'h2', function () {
            $(this).next('ul').toggle();
        });
    }

    bindEvents();

    // 3. 選單切換
    $('nav button').on('click', function () {
        $('nav ul').toggle();
    });
});
