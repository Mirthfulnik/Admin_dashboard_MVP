// автоскрытие логина, если уже авторизован
if (localStorage.getItem("lotus_auth") === "1") {
  const ov = document.getElementById("loginOverlay");
  if (ov) ov.style.display = "none";
} 

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

const eyeBtn = document.getElementById("togglePass");
const passInput = document.getElementById("loginPass");

if (eyeBtn && passInput) {
  eyeBtn.onclick = () => {
    const isHidden = passInput.type === "password";
    passInput.type = isHidden ? "text" : "password";
    eyeBtn.textContent = isHidden ? "🙈" : "👁";
  };
}
