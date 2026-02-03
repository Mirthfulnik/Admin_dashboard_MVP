const LOGIN_API = "https://functions.yandexcloud.net/d4eb50fp0hlkhk068v1e";

document.getElementById("loginBtn").onclick = async () => {
  const login = document.getElementById("loginUser").value.trim();
  const pass  = document.getElementById("loginPass").value.trim();

  const res = await fetch(LOGIN_API, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ login, password: pass })
  });

  if (res.ok) {
    document.getElementById("loginOverlay").style.display = "none";
    localStorage.setItem("lotus_auth", "1");
  } else {
    document.getElementById("loginError").innerText = "Неверный логин или пароль";
    document.getElementById("loginError").style.display = "block";
  }
};
