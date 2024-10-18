
(function () {
    "use strict";

    var messageBanner;

    // 每次加载新页面时都必须运行初始化函数。
    Office.initialize = function (reason) {

        Office.context.document.addHandlerAsync(Office.EventType.DocumentBeforeSave, handleBeforeSave, function (result) {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error(result.error.message);
            }
        });
    };

    // 这是处理 DocumentBeforeSave 事件的函数
    function handleBeforeSave(event) {
        // 在这里，你可以添加你想在保存文档前执行的代码
        console.log('文档即将保存');

        // 调用Word的另存为对话框
        Office.context.document.saveAsync(Office.SaveOptions.saveAs, function (result) {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error(result.error.message);
            } else {
                console.log('另存为对话框已打开。');
            }
        });

        // 如果你想阻止保存事件，可以设置 event.completed 为 false
        event.completed();
    }
})();
