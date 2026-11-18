
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport"
        content="width=device-width, initial-scale=1.0">
  <title>自动刷新时钟</title>

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
      font-size: 64px;
      font-weight: bold;
    }

    #date {
      font-size: 24px;
      margin-top: 20px;
      color: #aaa;
    }
  </style>
</head>

<body>
  <div id="clock"></div>
  <div id="date"></div>

  <script>
    // 获取电脑当前系统时间
    function updateTime() {
      const now = new Date();

      document.getElementById("clock").textContent =
        now.toLocaleTimeString("zh-CN", {
          hour12: false
        });

      document.getElementById("date").textContent =
        now.toLocaleDateString("zh-CN");
    }

    // 页面加载时显示当前时间
    updateTime();

    // 每隔60秒自动刷新整个页面
    setTimeout(function () {
      window.location.reload();
    }, 60 * 1000);
  </script>
</body>
</html>
