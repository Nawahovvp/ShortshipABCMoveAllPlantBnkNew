import { userUrl } from './config.js';

let users = [];
export let loggedUser = null;

export async function loadUsers() {
    try {
        const res = await fetch(userUrl);
        users = await res.json();
    } catch (e) {
        console.error(e);
        alert("ไม่สามารถโหลดข้อมูลผู้ใช้ได้");
    }
}

export function checkLogin(showLandingPageCallback) {
    const remembered = localStorage.getItem('rememberedUser');
    if (remembered) {
        loggedUser = JSON.parse(remembered);
        showUserMenu();
        if(showLandingPageCallback) showLandingPageCallback();
        return true;
    }
    const sessionUser = sessionStorage.getItem('loggedUser');
    if (sessionUser) {
        loggedUser = JSON.parse(sessionUser);
        showUserMenu();
        if(showLandingPageCallback) showLandingPageCallback();
        return true;
    }
    return false;
}

export function showLogin() {
    document.getElementById('loginModal').style.display = 'flex';
}

export function hideLogin() {
    document.getElementById('loginModal').style.display = 'none';
}

export function login(showLandingPageCallback) {
    const idUser = document.getElementById('idUserInput').value.trim();
    const password = document.getElementById('passwordInput').value.trim();
    const remember = document.getElementById('rememberMe').checked;
    const user = users.find(u => u.IDUser === idUser);
    if (user && password === idUser.slice(-4)) {
        loggedUser = { IDUser: user.IDUser, Name: user.Name || 'ไม่ระบุ' };
        if (remember) {
            localStorage.setItem('rememberedUser', JSON.stringify(loggedUser));
        } else {
            sessionStorage.setItem('loggedUser', JSON.stringify(loggedUser));
        }
        hideLogin();
        showUserMenu();
        if(showLandingPageCallback) showLandingPageCallback();
    } else {
        document.getElementById('loginError').textContent = 'IDUser หรือ Password ไม่ถูกต้อง';
    }
}

export function logout(showLoginCallback) {
    localStorage.removeItem('rememberedUser');
    sessionStorage.removeItem('loggedUser');
    loggedUser = null;
    if(showLoginCallback) showLoginCallback();
}

export function showUserMenu() {
    document.getElementById('menuBtn').style.display = 'block';
    document.getElementById('userID').textContent = `รหัส: ${loggedUser.IDUser}`;
    document.getElementById('userName').textContent = `ชื่อ: ${loggedUser.Name}`;
}