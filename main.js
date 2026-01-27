import { loadUsers, checkLogin, showLogin, login, logout, loggedUser } from './auth.js';
import { loadData, initABC, allData } from './abcApp.js';
import { loadShotShipData, initShotShip } from './shotShipApp.js';

document.addEventListener("DOMContentLoaded", function () {
    function handleLogout() {
        document.getElementById('userMenu').style.display = 'none';
        document.getElementById('menuBtn').style.display = 'none';
        document.getElementById('landingPage').style.display = 'none';
        document.getElementById('abcApp').style.display = 'none';
        document.getElementById('shotShipApp').style.display = 'none';
        logout(showLogin);
    }

    function showLandingPage() {
        document.getElementById('landingPage').style.display = 'flex';
        document.getElementById('abcApp').style.display = 'none';
        document.getElementById('shotShipApp').style.display = 'none';
        document.getElementById('appTitle').textContent = "ABC–Shortship Smart Ordering System";
        if (allData.length === 0) loadData(true);
    }

    function openABCApp() {
        document.getElementById('landingPage').style.display = 'none';
        document.getElementById('abcApp').style.display = 'block';
        document.getElementById('appTitle').textContent = "ABC–Shortship Smart Ordering System";
        loadData();
    }

    function openShotShipApp() {
        document.getElementById('landingPage').style.display = 'none';
        document.getElementById('abcApp').style.display = 'none';
        document.getElementById('shotShipApp').style.display = 'block';
        document.getElementById('appTitle').textContent = "ประสิทธิภาพในการจ่ายอะไหล่คลังสินค้า Sparepart ทั่วประเทศ (Short Ship)";
        loadShotShipData();
    }

    document.getElementById('menuBtn').addEventListener('click', () => {
        const menu = document.getElementById('userMenu');
        menu.style.display = menu.style.display === 'block' ? 'none' : 'block';
    });
    document.getElementById('logoutBtn').addEventListener('click', handleLogout);
    document.getElementById('loginBtn').addEventListener('click', () => login(showLandingPage));

    document.getElementById('btnMenuABC').addEventListener('click', openABCApp);
    document.getElementById('btnMenuShotShip').addEventListener('click', openShotShipApp);
    document.getElementById('btnBackToMenuABC').addEventListener('click', showLandingPage);
    document.getElementById('btnBackToMenuShotShip').addEventListener('click', showLandingPage);

    async function init() {
        await loadUsers();
        if (!checkLogin(showLandingPage)) showLogin();
        initABC();
        initShotShip();
    }
    init();
});