<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Spark Automation | IoT System</title>
    
    <!-- External Libraries for Excel Export -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js"></script>
    
    <style>
        :root {
            --spark-blue: #0056b3;
            --spark-dark: #1a1a1a;
            --spark-light: #f4f7f6;
            --success: #28a745;
            --danger: #dc3545;
        }

        body { font-family: 'Segoe UI', sans-serif; margin: 0; background: var(--spark-light); }
        .page { display: none; min-height: 100vh; }
        .active { display: block; }

        /* Login Page Styling */
        #loginPage { 
            background: linear-gradient(135deg, var(--spark-dark) 0%, var(--spark-blue) 100%); 
            display: flex; align-items: center; justify-content: center; 
        }
        .login-card { 
            background: white; padding: 40px; border-radius: 15px; 
            width: 100%; max-width: 350px; text-align: center; 
            box-shadow: 0 10px 25px rgba(0,0,0,0.3);
        }
        .login-card input { 
            width: 100%; padding: 12px; margin: 10px 0; 
            border: 1px solid #ddd; border-radius: 5px; box-sizing: border-box; 
        }

        /* Spark Header Block (Top 1/4th) */
        .spark-header { 
            background: var(--spark-dark); color: white; 
            height: 25vh; border-bottom: 8px solid var(--spark-blue);
            display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center;
        }
        .spark-header h1 { margin: 0; font-size: 2.5rem; letter-spacing: 2px; }
        .spark-header p { margin: 5px 0; font-size: 0.9rem; opacity: 0.9; }

        /* Navigation Cards */
        .home-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 30px; padding: 40px; max-width: 800px; margin: auto; }
        .card { 
            background: white; padding: 50px; border-radius: 20px; text-align: center; 
            cursor: pointer; box-shadow: 0 4px 15px rgba(0,0,0,0.1); transition: 0.3s;
            border: 2px solid transparent;
        }
        .card:hover { transform: translateY(-10px); border-color: var(--spark-blue); }
        .card-icon { font-size: 60px; margin-bottom: 15px; }

        /* Container for Pages */
        .container { max-width: 800px; margin: 20px auto; background: white; padding: 30px; border-radius: 15px; }
        .top-bar { background: #f0f0f0; padding: 10px 20px; display: flex; justify-content: space-between; border-radius: 8px; margin-bottom: 20px; }
        
        /* Buttons */
        .btn { padding: 12px 25px; border: none; border-radius: 5px; cursor: pointer; font-weight: bold; }
        .btn-on { background: var(--success); color: white; margin-right: 10px; }
        .btn-off { background: var(--danger); color: white; }
        .btn-main { background: var(--spark-blue); color: white; width: 100%; margin-top: 15px; }

        /* Data Table */
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { padding: 12px; border: 1px solid #ddd; text-align: left; }
        th { background: #f8f9fa; }
    </style>
</head>
<body>

    <!-- PAGE 1: LOGIN -->
    <div id="loginPage" class="page active">
        <div class="login-card">
            <h2 style="color: var(--spark-blue);">SPARK DNA LOGIN</h2>
            <p style="font-size: 0.8rem; color: #666;">Enter credentials to access IoT Portal</p>
            <input type="email" id="userEmail" placeholder="Email (24e615@psgtech.ac.in)">
            <input type="password" id="userPass" placeholder="Password (123456789)">
            <button class="btn btn-main" onclick="handleLogin()">ACCESS SYSTEM</button>
        </div>
    </div>

    <!-- PAGE 2: HOME -->
    <div id="homePage" class="page">
        <div class="spark-header">
            <h1>SPARK AUTOMATION</h1>
            <p>Industrial Automation & IoT Solutions | www.sparkdna.co.in</p>
            <p>Contact: +91 99942 34433 | Email: contact@sparkdna.co.in</p>
            <p>PLC Panels • SCADA • Machine Automation</p>
        </div>
        <div class="home-grid">
            <div class="card" onclick="showPage('controlPage')">
                <div class="card-icon">🎮</div>
                <h2>CONTROL</h2>
            </div>
            <div class="card" onclick="showPage('dataPage')">
                <div class="card-icon">📊</div>
                <h2>DATA</h2>
            </div>
        </div>
    </div>

    <!-- PAGE 3: CONTROL -->
    <div id="controlPage" class="page">
        <div class="container">
            <div class="top-bar">
                <strong>SPARK AUTOMATION</strong>
                <button onclick="showPage('homePage')" class="btn" style="padding: 5px 15px;">🏠 Home</button>
            </div>
            <h2 style="text-align:center">Device Control Center</h2>
            
            <div style="text-align: center; margin: 30px 0;">
                <h3>Status: <span id="ledStatusTxt" style="color: gray;">...</span></h3>
                <button class="btn btn-on" onclick="sendManualCommand(1)">TURN ON</button>
                <button class="btn btn-off" onclick="sendManualCommand(0)">TURN OFF</button>
            </div>

            <div style="border-top: 1px solid #eee; padding-top: 20px;">
                <h4>Operation Mode: 
                    <select id="modeSelector" onchange="toggleMode()">
                        <option value="manual">Manual Mode</option>
                        <option value="auto">Auto Mode</option>
                    </select>
                </h4>
                
                <div id="autoSection" style="display:none; background:#fdfdfd; padding:20px; border: 1px dashed #ccc;">
                    <label>Start Date/Time:</label><br>
                    <input type="datetime-local" id="startTime" style="width:100%; padding:8px; margin-bottom:15px;">
                    <label>End Date/Time:</label><br>
                    <input type="datetime-local" id="endTime" style="width:100%; padding:8px; margin-bottom:15px;">
                    <label>Interval Type:</label>
                    <select id="intervalType" style="padding:8px;">
                        <option value="sec">Second wise</option>
                        <option value="min">Minute wise</option>
                        <option value="hour">Hour wise</option>
                    </select>
                    <button class="btn btn-main" onclick="applyAuto()">Save Auto Schedule</button>
                </div>
            </div>
        </div>
    </div>

    <!-- PAGE 4: DATA -->
    <div id="dataPage" class="page">
        <div class="container">
            <div class="top-bar">
                <strong>SPARK AUTOMATION</strong>
                <button onclick="showPage('homePage')" class="btn" style="padding: 5px 15px;">🏠 Home</button>
            </div>
            <h2>Data and History</h2>
            <div style="background: #f4f4f4; padding: 15px; border-radius: 8px;">
                Filter Date: <input type="date" id="filterFrom"> to <input type="date" id="filterTo">
                <button class="btn btn-main" style="width: auto; margin: 0 10px;" onclick="loadHistory()">Apply</button>
            </div>
            <table id="historyTable">
                <thead>
                    <tr><th>Date</th><th>Time</th><th>Status</th><th>Mode</th></tr>
                </thead>
                <tbody id="historyBody"></tbody>
            </table>
            <button class="btn btn-main" onclick="downloadExcel()">📥 Download Excel Report</button>
        </div>
    </div>

    <script type="module">
        import { initializeApp } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js";
        import { getDatabase, ref, set, onValue, push, update } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-database.js";
        import { getAuth, signInWithEmailAndPassword } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-auth.js";

        const firebaseConfig = {
            apiKey: "AIzaSyBO7oSsNJd_Hg60p95t-L71Vi_0iXRFS5M",
            authDomain: "spark-automation-90cb1.firebaseapp.com",
            databaseURL: "https://spark-automation-90cb1-default-rtdb.asia-southeast1.firebasedatabase.app",
            projectId: "spark-automation-90cb1",
            storageBucket: "spark-automation-90cb1.firebasestorage.app",
            messagingSenderId: "53319549438",
            appId: "1:53319549438:web:6cc6b39da070892df980af"
        };

        const app = initializeApp(firebaseConfig);
        const db = getDatabase(app);
        const auth = getAuth(app);

        // --- NAVIGATION ---
        window.showPage = (id) => {
            document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
            document.getElementById(id).classList.add('active');
        };

        // --- AUTHENTICATION ---
        window.handleLogin = () => {
            const email = document.getElementById('userEmail').value;
            const pass = document.getElementById('userPass').value;
            // Using your specific credentials
            signInWithEmailAndPassword(auth, email, pass)
                .then(() => showPage('homePage'))
                .catch(err => alert("Login Failed: Incorrect Email or Password"));
        };

        // --- CONTROL ---
        window.toggleMode = () => {
            const m = document.getElementById('modeSelector').value;
            update(ref(db, 'CONTROL'), { mode: m });
            document.getElementById('autoSection').style.display = (m === 'auto') ? 'block' : 'none';
        };

        window.sendManualCommand = (val) => {
            update(ref(db, 'CONTROL'), { led_status: val, mode: 'manual' });
            saveLog(val, "Manual");
        };

        window.applyAuto = () => {
            const start = document.getElementById('startTime').value;
            const end = document.getElementById('endTime').value;
            const inter = document.getElementById('intervalType').value;
            update(ref(db, 'CONTROL/auto_settings'), { 
                start_time: start, 
                end_time: end, 
                interval: inter 
            });
            alert("Auto Logic Applied!");
        };

        // Status Listener
        onValue(ref(db, 'CONTROL/led_status'), (snap) => {
            const val = snap.val();
            const el = document.getElementById('ledStatusTxt');
            el.innerText = (val === 1) ? "DEVICE IS ON" : "DEVICE IS OFF";
            el.style.color = (val === 1) ? "var(--success)" : "var(--danger)";
        });

        // --- DATA LOGGING ---
        function saveLog(status, mode) {
            const now = new Date();
            push(ref(db, 'HISTORY'), {
                date: now.toISOString().split('T')[0],
                time: now.toLocaleTimeString(),
                status: (status === 1) ? "ON" : "OFF",
                mode: mode
            });
        }

        window.loadHistory = () => {
            onValue(ref(db, 'HISTORY'), (snap) => {
                const logs = snap.val();
                const body = document.getElementById('historyBody');
                body.innerHTML = "";
                for (let id in logs) {
                    body.innerHTML += `<tr>
                        <td>${logs[id].date}</td>
                        <td>${logs[id].time}</td>
                        <td>${logs[id].status}</td>
                        <td>${logs[id].mode}</td>
                    </tr>`;
                }
            });
        };

        window.downloadExcel = () => {
            const table = document.getElementById("historyTable");
            const wb = XLSX.utils.table_to_book(table);
            XLSX.writeFile(wb, "Spark_Automation_Report.xlsx");
        };

        // Browser Back Support
        window.onpopstate = () => showPage('homePage');
        loadHistory(); // Load data on start

    </script>
</body>
</html>
