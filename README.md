
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport"
        content="width=device-width, initial-scale=1.0">
  <title>自动刷新时间页面</title>

  <style>
    body {
      margin: 0;
      height: 100vh;
      display: flex;
      flex-direction: column;
      justify-content: center;
      align-items: center;
      background: #121212;
      color: white;
      font-family: Arial, sans-serif;
    }

    #clock {
      font-size: 60px;
      font-weight: bold;
    }

    #date {
      font-size: 24px;
      margin: 20px 0;
      color: #aaa;
    }

    #timeInput {
      font-size: 24px;
      padding: 12px;
      width: 260px;
      text-align: center;
      border: 2px solid #4caf50;
      border-radius: 8px;
      outline: none;
    }

    #status {
      margin-top: 20px;
      color: #aaa;
    }
  </style>
</head>

<body>

  <div id="clock"></div>
  <div id="date"></div>

  <input
    type="text"
    id="timeInput"
    placeholder="当前时间"
    autocomplete="off"
  >

  <div id="status">
    页面每60秒自动刷新
  </div>

  <script>
    function initializePage() {

      // 获取电脑当前系统时间
      const now = new Date();

      const time = now.toLocaleTimeString(
        "zh-CN",
        { hour12: false }
      );

      const date = now.toLocaleDateString("zh-CN");

      // 显示日期和时间
      document.getElementById("clock").textContent =
        time;

      document.getElementById("date").textContent =
        date;

      // 获取输入框
      const input =
        document.getElementById("timeInput");

      // 自动聚焦输入框
      input.focus();

      // 模拟用户输入当前时间
      const setter = Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        "value"
      ).set;

      setter.call(input, time);

      // 触发 input 事件
      input.dispatchEvent(
        new Event("input", { bubbles: true })
      );

      // 触发 change 事件
      input.dispatchEvent(
        new Event("change", { bubbles: true })
      );

      // 光标移动到末尾
      input.setSelectionRange(
        input.value.length,
        input.value.length
      );
    }

    // 页面加载完成后执行
    window.addEventListener("load", () => {

      initializePage();

      // 每60秒真正刷新整个HTML页面
      setTimeout(() => {
        window.location.reload();
      }, 60000);

    });
  </script>

</body>
</html>
