// JavaScript 代碼將在這裡添加
window.onload = function () {
    var articles = {
        "短歌行": [
            "十五音切語",
            "台羅拼音",
            "台羅改良式",
            "方音符號注音",
            "白話字拼音",
            "閩拼標音",
            "雙排注音"
        ]
    };

    var articlesDiv = document.getElementById('articles');

    // 遍歷每個文章
    for (var article in articles) {
        if (articles.hasOwnProperty(article)) {
            var articleTitle = document.createElement('h2');
            articleTitle.textContent = article;
            articlesDiv.appendChild(articleTitle);

            var phoneticMethods = articles[article];
            var phoneticList = document.createElement('ul');

            // 遍歷每種注音方式
            for (var i = 0; i < phoneticMethods.length; i++) {
                var phoneticMethod = phoneticMethods[i];
                var listItem = document.createElement('li');
                var link = document.createElement('a');
                link.href = article + '_' + phoneticMethod + '.html';
                link.textContent = phoneticMethod;
                listItem.appendChild(link);
                phoneticList.appendChild(listItem);
            }

            articlesDiv.appendChild(phoneticList);
        }
    }
};
