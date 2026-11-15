<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>منظومة إدارة - مكتب تدريب طاقم الضيافة الجوية - الخطوط الجوية الليبية</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Fonts: Cairo & Tajawal -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&family=Tajawal:wght@400;500;700&display=swap" rel="stylesheet">
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- SheetJS for Excel Export -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

    <script>
        tailwind.config = {
            theme: {
                extend: {
                    colors: {
                        libyan: {
                            blue: '#004b87',
                            sky: '#0066b2',
                            darkBlue: '#1a1814',
                            gold: '#C5A059',
                            goldLight: '#E2C787',
                            darkGold: '#997A3B',
                            lightBg: '#F9F8F6'
                        }
                    },
                    fontFamily: {
                        cairo: ['Cairo', 'sans-serif'],
                        tajawal: ['Tajawal', 'sans-serif']
                    }
                }
            }
        }
    </script>
    <style>
        body { font-family: 'Cairo', sans-serif; }
        .tab-btn.active {
            background-color: white;
            color: #997A3B;
            font-weight: 700;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }
    </style>
</head>
<body class="bg-gray-100 min-h-screen text-gray-800">

    <!-- شاشة تسجيل الدخول السرية -->
    <div id="login-screen" class="fixed inset-0 bg-libyan-darkBlue z-50 flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl max-w-md w-full p-8 shadow-2xl text-center space-y-6 border-t-4 border-libyan-gold">
            <div class="w-16 h-16 bg-amber-50 text-libyan-gold rounded-full flex items-center justify-center mx-auto text-2xl shadow-inner">
                <i class="fa-solid fa-lock"></i>
            </div>
            <div>
                <h2 class="text-2xl font-bold text-gray-800">تسجيل دخول المنظومة</h2>
                <p class="text-xs text-gray-500 mt-1">الخطوط الجوية الليبية - مكتب تدريب الضيافة</p>
            </div>
            <div class="space-y-3 text-right">
                <label class="block text-xs font-semibold text-gray-700">أدخل الكود السري للمتابعة:</label>
                <input type="password" id="secret-pass" placeholder="أدخل الكود هنا..." class="w-full p-3 border border-gray-300 rounded-xl text-center tracking-widest text-lg font-bold focus:outline-none focus:border-libyan-gold">
                <p id="login-error" class="text-red-500 text-xs text-center hidden">الكود السري غير صحيح، حاول مرة أخرى.</p>
            </div>
            <button onclick="verifyLogin()" class="w-full bg-libyan-gold hover:bg-libyan-darkGold text-white font-bold py-3 rounded-xl transition shadow-md">
                دخول المنظومة ✈️
            </button>
        </div>
    </div>

    <!-- محتوى المنظومة الأساسي -->
    <div id="app-container" class="hidden pb-12">
        <header class="bg-gradient-to-r from-libyan-darkBlue via-zinc-900 to-libyan-darkGold text-white shadow-lg sticky top-0 z-40 border-b border-libyan-gold/30">
            <div class="max-w-7xl mx-auto px-4 py-4 flex justify-between items-center">
                <div class="flex items-center space-x-3 space-x-reverse">
                    <div class="bg-white/10 p-1 rounded-full w-12 h-12 flex items-center justify-center shadow-md overflow-hidden border border-libyan-gold shrink-0 backdrop-blur-sm">
                        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 426 406" class="w-full h-full">
                            <g id="surface1">
                                <path style="stroke:none;fill-rule:evenodd;fill:#D9252A;fill-opacity:1;" d="M 72.742188 168.730469 C 79.222656 161.207031 94.4375 161.945313 100.921875 168.730469 C 100.492188 169.753906 96.25 172.296875 94.996094 172.867188 C 86.382813 176.808594 73.71875 173.742188 72.742188 168.730469 Z M 106.289063 164.125 L 106.285156 130.140625 C 105.242188 130.402344 89.578125 145.273438 88.449219 146.3125 C 82.566406 151.722656 76.5 157.25 70.742188 162.644531 C 64.820313 168.195313 58.949219 173.5625 53.046875 178.980469 L 37.5 193.328125 C 36.515625 194.246094 35.625 194.890625 35.332031 195.738281 C 43.957031 195.464844 52.980469 195.707031 61.652344 195.707031 L 101.207031 195.671875 C 108.738281 195.671875 106.816406 196.523438 112.488281 193.578125 C 118.066406 190.683594 124.109375 187.453125 129.894531 184.554688 C 132.796875 183.097656 198.847656 148.511719 199.640625 148.652344 C 200.578125 149.144531 201.675781 150.132813 202.566406 150.863281 L 211.464844 157.828125 C 215.535156 160.886719 219.089844 163.867188 223.203125 167.015625 C 227.0625 169.96875 231.238281 173.472656 234.941406 176.199219 C 236.230469 177.144531 246.273438 184.941406 246.578125 185.554688 L 202.773438 185.582031 L 202.761719 166.746094 C 201.78125 166.597656 189.136719 173.457031 187.214844 174.464844 C 180.882813 177.789063 140.84375 198.429688 140.539063 198.851563 C 141.011719 199.648438 143.339844 200.582031 144.136719 201.019531 C 145.410156 201.722656 146.570313 202.386719 147.992188 203.132813 C 150.859375 204.636719 152.929688 205.835938 155.578125 207.265625 L 178.339844 219.671875 C 181.015625 221.132813 183.144531 222.25 185.96875 223.8125 C 187.988281 224.929688 200.410156 231.902344 201.5625 232.058594 C 201.554688 226.351563 201.390625 218.871094 201.566406 213.246094 L 280.824219 213.164063 C 282.222656 213.164063 282.480469 213.507813 283.238281 214.144531 C 289.8125 219.671875 298.699219 226.121094 305.480469 231.703125 L 317.570313 241.214844 C 323.730469 246.152344 332.1875 252.363281 337.914063 257.214844 L 343.957031 261.972656 C 344.980469 262.734375 344.84375 262.949219 346.601563 262.949219 L 389.378906 262.949219 C 391.433594 262.949219 390.902344 263.035156 390.902344 257.492188 C 390.902344 252.074219 391.21875 239.636719 390.871094 235.753906 C 389.324219 235.285156 364.40625 235.578125 359.246094 235.578125 C 353.03125 235.578125 354.660156 236.097656 350.988281 233.097656 C 346.933594 229.789063 343.355469 226.703125 339.371094 223.328125 L 330.738281 216.097656 C 330.167969 215.621094 328.148438 214.125 328.015625 213.414063 C 338.460938 213.097656 349.433594 213.40625 359.9375 213.40625 L 389.863281 213.414063 C 390.929688 213.195313 390.597656 213.238281 390.597656 211.960938 C 390.597656 206.335938 390.839844 190.085938 390.503906 185.628906 C 385.628906 185.257813 379.386719 185.59375 374.429688 185.59375 C 347.417969 185.59375 319.671875 185.304688 292.78125 185.53125 C 289.710938 185.558594 290.917969 185.53125 287.382813 182.867188 L 273.621094 172.007813 C 267.746094 167.308594 261.636719 162.671875 255.5 157.710938 L 218.988281 128.910156 C 214.472656 125.429688 204.574219 117.179688 200.789063 114.535156 L 189.050781 120.652344 C 165.914063 133.160156 141.523438 145.550781 118.34375 158.015625 C 116.585938 158.957031 107.328125 164.078125 106.289063 164.125 "/>
                                <path style="stroke:none;fill-rule:evenodd;fill:#C5A059;fill-opacity:1;" d="M 279.351563 21.179688 C 281.230469 21.921875 285.089844 22.996094 288.414063 24.316406 C 303.480469 30.300781 315.78125 36.476563 328.980469 45.394531 C 334.078125 48.835938 338.75 52.207031 343.160156 55.8125 C 360.652344 70.117188 376.359375 87.132813 388.964844 108.648438 C 401.234375 129.589844 410.503906 154.132813 413.621094 182.355469 C 416.984375 212.835938 412.300781 243.363281 402.789063 268.15625 C 398.269531 279.933594 392.753906 290.976563 386.460938 301.03125 C 373.605469 321.554688 358.695313 337.546875 339.574219 352.566406 C 330.074219 360.027344 319.832031 366.164063 309.386719 371.800781 C 294.039063 380.082031 272.3125 387.761719 254.742188 391.015625 C 238.320313 394.058594 226.5625 395.289063 209.53125 395.117188 C 188.496094 394.910156 165.226563 390.941406 146.144531 384.363281 C 108.039063 371.230469 76.253906 349.378906 51.902344 318.757813 C 33.796875 295.996094 22.640625 272.492188 15.703125 243.882813 C 14.011719 236.910156 13.175781 230.316406 12.183594 222.8125 C 10.277344 208.453125 11.039063 190.171875 13.015625 176.234375 C 17.492188 144.667969 30.589844 114.382813 49.6875 89.625 C 53.265625 84.984375 57.070313 80.621094 60.921875 76.246094 C 68.066406 68.128906 77.574219 59.625 86.617188 52.738281 C 99.730469 42.75 117.707031 31.929688 133.695313 25.777344 C 155.777344 17.277344 169.234375 13.96875 193.152344 11.332031 C 207.644531 9.734375 225.125 10.136719 239.5625 12.039063 C 246.621094 12.96875 253.628906 14.140625 260.472656 15.742188 C 266.472656 17.148438 273.667969 18.925781 279.351563 21.179688 Z M 203.894531 0.464844 C 199.972656 0.902344 195.539063 0.5625 191.652344 1.164063 C 189.753906 1.460938 187.679688 1.476563 185.664063 1.726563 C 162.9375 4.542969 140.628906 11.105469 120.097656 20.457031 C 74.636719 41.15625 37.105469 78.714844 16.792969 124.582031 C 4.660156 151.980469 -0.492188 180.8125 0.457031 211.964844 C 0.621094 217.355469 2.234375 229.984375 3.078125 235.113281 C 6.578125 256.390625 13.167969 274.460938 22.761719 293.375 C 44.703125 336.65625 87.886719 373.90625 133.03125 390.570313 C 143.398438 394.398438 161.898438 400.414063 173.425781 401.894531 C 177.199219 402.382813 180.753906 403.199219 184.691406 403.578125 C 188.59375 403.953125 192.570313 404.53125 196.644531 404.761719 C 204.984375 405.234375 213.578125 405.351563 222.179688 405.148438 C 250.085938 404.488281 280.695313 396.640625 305.863281 385.0625 C 329.523438 374.175781 350.328125 358.898438 366.902344 342.40625 C 380.945313 328.425781 390.269531 315.542969 400.441406 298.214844 C 413.273438 276.359375 423.308594 244.1875 424.699219 218.625 C 424.796875 216.828125 425.121094 214.667969 425.34375 213.039063 C 425.671875 210.589844 425.523438 196.054688 425.414063 193.554688 L 424.871094 187.742188 C 422.742188 166.359375 419.585938 150.328125 411.3125 130.039063 C 408.773438 123.816406 405.933594 117.722656 403.089844 112.25 C 391.085938 89.164063 376.328125 71.519531 357.40625 54.214844 C 340.1875 38.464844 314.769531 23.25 292.894531 15.011719 C 279.4375 9.945313 275.683594 9.042969 263.34375 5.914063 C 260.082031 5.089844 255.828125 4.195313 252.511719 3.667969 C 240.109375 1.699219 229.660156 0.375 216.890625 0.296875 C 212.695313 0.273438 207.898438 0.0195313 203.894531 0.464844 "/>
                                <path style="stroke:none;fill-rule:evenodd;fill:#D9252A;fill-opacity:1;" d="M 44.238281 247.605469 L 74.917969 264.707031 C 79.722656 267.304688 85.585938 270.691406 90.246094 273.265625 C 92.796875 274.671875 95.570313 276.160156 97.945313 277.539063 C 99.21875 278.277344 100.605469 278.992188 101.777344 279.679688 C 102.574219 280.144531 104.867188 281.667969 105.691406 281.546875 L 105.6875 262.675781 L 294.164063 263.058594 C 296.1875 263.058594 295.546875 262.757813 295.546875 257.03125 L 295.675781 234.875 C 283.851563 234.589844 269.820313 234.90625 257.824219 234.9375 L 133.867188 235.285156 C 130.078125 235.285156 107.804688 235.511719 106.101563 235.160156 L 106.089844 215.117188 C 104.71875 215.269531 56.792969 240.671875 51.832031 243.289063 L 45.03125 246.832031 C 44.101563 247.371094 44.457031 246.933594 44.238281 247.605469 "/>
                                <path style="stroke:none;fill-rule:evenodd;fill:#1a1814;fill-opacity:1;" d="M 31.917969 194.089844 C 33.363281 193.859375 35.601563 191.042969 36.765625 189.839844 L 50.425781 176.359375 C 56.980469 169.785156 76.398438 151.648438 83.335938 146.257813 C 92.550781 139.097656 98.90625 133.453125 109.214844 126.601563 C 112.960938 124.109375 116.222656 121.90625 120.113281 119.574219 C 131.164063 112.960938 143.992188 105.996094 156.046875 101.210938 C 175.496094 93.484375 183.832031 90.722656 203.136719 85.085938 C 204.355469 84.730469 205.542969 84.496094 206.78125 84.101563 C 208.117188 83.679688 209.699219 83.851563 209.675781 82.492188 C 207.738281 81.214844 184.117188 88.222656 180.78125 89.152344 C 168.988281 92.449219 152.457031 98.832031 141.046875 103.730469 C 136.871094 105.519531 132.636719 107.5 128.65625 109.488281 C 120.820313 113.402344 112.792969 117.871094 105.484375 122.648438 C 90.40625 132.507813 65.777344 154.019531 54.289063 166.199219 C 49.898438 170.851563 43.082031 178.695313 39.535156 183.410156 C 38.242188 185.128906 32.222656 192.636719 31.917969 194.089844 "/>
                                <path style="stroke:none;fill-rule:evenodd;fill:#D9252A;fill-opacity:1;" d="M 75.285156 168.960938 C 76.113281 172.179688 86.738281 174.0625 93.074219 171.613281 C 94.386719 171.105469 97.125 169.847656 97.777344 168.730469 C 95.042969 165.222656 80.761719 163.136719 75.285156 168.960938 "/>
                            </g>
                        </svg>
                    </div>
                    <div>
                        <h1 class="text-xl md:text-2xl font-bold font-cairo tracking-wide text-libyan-goldLight">الخطوط الجوية الليبية</h1>
                        <p class="text-xs md:text-sm text-gray-300 font-tajawal">منظومة إدارية - مكتب تدريب طاقم الضيافة الجوية</p>
                    </div>
                </div>
                <div class="flex items-center space-x-2 space-x-reverse">
                    <button onclick="lockSystem()" class="p-2 bg-white/10 hover:bg-white/20 rounded-lg transition text-libyan-goldLight" title="قفل المنظومة">
                        <i class="fa-solid fa-lock text-lg"></i>
                    </button>
                </div>
            </div>
        </header>

        <main class="max-w-7xl mx-auto px-4 mt-6">
            <!-- Stats Grid -->
            <div class="grid grid-cols-2 md:grid-cols-4 gap-3 mb-6">
                <div class="bg-white p-3 rounded-xl shadow-sm border border-gray-100 flex items-center justify-between border-r-4 border-r-libyan-gold">
                    <div>
                        <p class="text-[11px] font-semibold text-gray-500 mb-0.5">إجمال الضيافة</p>
                        <p id="stat-total" class="text-xl font-bold text-gray-800">0</p>
                    </div>
                    <div class="w-8 h-8 bg-amber-50 text-libyan-gold rounded-lg flex items-center justify-center text-sm">
                        <i class="fa-solid fa-users"></i>
                    </div>
                </div>

                <!-- بطاقة عدد المضيفين مع تفصيل النشطين والمنتهين -->
                <div class="bg-white p-3 rounded-xl shadow-sm border border-gray-100 border-r-4 border-r-blue-600">
                    <div class="flex items-center justify-between">
                        <div>
                            <p class="text-[11px] font-semibold text-gray-500 mb-0.5">عدد المضيفين (الإجمالي)</p>
                            <p id="stat-males" class="text-xl font-bold text-blue-600">0</p>
                        </div>
                        <div class="w-8 h-8 bg-blue-50 text-blue-600 rounded-lg flex items-center justify-center text-sm">
                            <i class="fa-solid fa-mars"></i>
                        </div>
                    </div>
                    <div class="flex justify-between text-[11px] mt-2 pt-2 border-t border-gray-100 text-gray-600 font-semibold">
                        <span>النشطين: <strong id="stat-males-active" class="text-green-600">0</strong></span>
                        <span>المنتهين: <strong id="stat-males-expired" class="text-red-600">0</strong></span>
                    </div>
                </div>

                <!-- بطاقة عدد المضيفات مع تفصيل النشطين والمنتهين -->
                <div class="bg-white p-3 rounded-xl shadow-sm border border-gray-100 border-r-4 border-r-pink-500">
                    <div class="flex items-center justify-between">
                        <div>
                            <p class="text-[11px] font-semibold text-gray-500 mb-0.5">عدد المضيفات (الإجمالي)</p>
                            <p id="stat-females" class="text-xl font-bold text-pink-500">0</p>
                        </div>
                        <div class="w-8 h-8 bg-pink-50 text-pink-500 rounded-lg flex items-center justify-center text-sm">
                            <i class="fa-solid fa-venus"></i>
                        </div>
                    </div>
                    <div class="flex justify-between text-[11px] mt-2 pt-2 border-t border-gray-100 text-gray-600 font-semibold">
                        <span>النشطين: <strong id="stat-females-active" class="text-green-600">0</strong></span>
                        <span>المنتهين: <strong id="stat-females-expired" class="text-red-600">0</strong></span>
                    </div>
                </div>

                <!-- بطاقة رؤساء الأطقم مع تفصيل النشطين والمنتهين -->
                <div class="bg-white p-3 rounded-xl shadow-sm border border-gray-100 border-r-4 border-r-purple-600">
                    <div class="flex items-center justify-between">
                        <div>
                            <p class="text-[11px] font-semibold text-gray-500 mb-0.5">رؤساء الأطقم (الإجمالي)</p>
                            <p id="stat-leaders" class="text-xl font-bold text-purple-600">0</p>
                        </div>
                        <div class="w-8 h-8 bg-purple-50 text-purple-600 rounded-lg flex items-center justify-center text-sm">
                            <i class="fa-solid fa-user-tie"></i>
                        </div>
                    </div>
                    <div class="flex justify-between text-[11px] mt-2 pt-2 border-t border-gray-100 text-gray-600 font-semibold">
                        <span>النشطين: <strong id="stat-leaders-active" class="text-green-600">0</strong></span>
                        <span>المنتهين: <strong id="stat-leaders-expired" class="text-red-600">0</strong></span>
                    </div>
                </div>

                <div class="bg-white p-3 rounded-xl shadow-sm border border-gray-100 flex items-center justify-between border-r-4 border-r-cyan-600">
                    <div>
                        <p class="text-[11px] font-semibold text-gray-500 mb-0.5">في إجازة</p>
                        <p id="stat-leaves" class="text-xl font-bold text-cyan-600">0</p>
                    </div>
                    <div class="w-8 h-8 bg-cyan-50 text-cyan-600 rounded-lg flex items-center justify-center text-sm">
                        <i class="fa-solid fa-umbrella-beach"></i>
                    </div>
                </div>
                <div class="bg-white p-3 rounded-xl shadow-sm border border-gray-100 flex items-center justify-between border-r-4 border-r-green-600">
                    <div>
                        <p class="text-[11px] font-semibold text-gray-500 mb-0.5">ساري المفعول</p>
                        <p id="stat-valid" class="text-xl font-bold text-green-600">0</p>
                    </div>
                    <div class="w-8 h-8 bg-green-50 text-green-600 rounded-lg flex items-center justify-center text-sm">
                        <i class="fa-solid fa-circle-check"></i>
                    </div>
                </div>
                <div class="bg-white p-3 rounded-xl shadow-sm border border-gray-100 flex items-center justify-between border-r-4 border-r-amber-500">
                    <div>
                        <p class="text-[11px] font-semibold text-gray-500 mb-0.5">قريب الانتهاء</p>
                        <p id="stat-warning" class="text-xl font-bold text-amber-500">0</p>
                    </div>
                    <div class="w-8 h-8 bg-amber-50 text-amber-500 rounded-lg flex items-center justify-center text-sm">
                        <i class="fa-solid fa-triangle-exclamation"></i>
                    </div>
                </div>
                <div class="bg-white p-3 rounded-xl shadow-sm border border-gray-100 flex items-center justify-between border-r-4 border-r-red-600">
                    <div>
                        <p class="text-[11px] font-semibold text-gray-500 mb-0.5">منتهية الصلاحية</p>
                        <p id="stat-expired" class="text-xl font-bold text-red-600">0</p>
                    </div>
                    <div class="w-8 h-8 bg-red-50 text-red-600 rounded-lg flex items-center justify-center text-sm">
                        <i class="fa-solid fa-circle-xmark"></i>
                    </div>
                </div>
            </div>

            <!-- Main Navigation Tabs -->
            <div class="bg-amber-100/60 p-1.5 rounded-xl flex space-x-1 space-x-reverse mb-6 text-sm border border-libyan-gold/20">
                <button onclick="switchTab('crew')" id="tab-crew" class="tab-btn active flex-1 py-2.5 rounded-lg text-center transition flex items-center justify-center space-x-2 space-x-reverse">
                    <i class="fa-solid fa-id-card"></i>
                    <span>دليل وسجلات المضيفين</span>
                </button>
                <button onclick="switchTab('courses')" id="tab-courses" class="tab-btn flex-1 py-2.5 rounded-lg text-center transition flex items-center justify-center space-x-2 space-x-reverse">
                    <i class="fa-solid fa-graduation-cap"></i>
                    <span>متابعة الدورات التفصيلية</span>
                </button>
                <button onclick="switchTab('reports')" id="tab-reports" class="tab-btn flex-1 py-2.5 rounded-lg text-center transition flex items-center justify-center space-x-2 space-x-reverse">
                    <i class="fa-solid fa-print"></i>
                    <span>التقارير والتصدير</span>
                </button>
            </div>

            <!-- Section: Crew Management -->
            <div id="section-crew" class="space-y-4">
                <div class="bg-white p-4 rounded-xl shadow-sm border border-gray-100 space-y-3">
                    <div class="flex flex-col md:flex-row gap-3">
                        <div class="relative flex-1">
                            <i class="fa-solid fa-magnifying-glass absolute right-3 top-3.5 text-gray-400"></i>
                            <input type="text" id="search-input" oninput="renderCrewList()" placeholder="ابحث باسم المضيف، رقم الملف (F-101)، أو الرقم الوظيفي..." class="w-full pr-10 pl-4 py-2.5 bg-gray-50 border border-gray-200 rounded-lg text-sm focus:outline-none focus:border-libyan-gold transition">
                        </div>
                        <select id="filter-aircraft" onchange="renderCrewList()" class="bg-gray-50 border border-gray-200 rounded-lg px-3 py-2.5 text-sm focus:outline-none focus:border-libyan-gold">
                            <option value="ALL">جميع الطرازات (A320 & A330)</option>
                            <option value="A320">طراز A320 فقط</option>
                            <option value="A330">طراز A330 فقط</option>
                        </select>
                        <button onclick="openAddCrewModal()" class="bg-libyan-gold hover:bg-libyan-darkGold text-white font-semibold px-4 py-2.5 rounded-lg text-sm transition flex items-center justify-center space-x-2 space-x-reverse shadow-sm">
                            <i class="fa-solid fa-user-plus"></i>
                            <span>إضافة مضيف جديد</span>
                        </button>
                    </div>

                    <!-- أزرار الفلترة الرئيسية -->
                    <div class="flex flex-wrap gap-2 pt-2 border-t border-gray-100 text-xs">
                        <button onclick="setFilterStatus('ALL')" id="filter-btn-ALL" class="px-3 py-1.5 rounded-lg font-bold bg-libyan-darkBlue text-white">عرض الكل</button>
                        <button onclick="setFilterStatus('VALID')" id="filter-btn-VALID" class="px-3 py-1.5 rounded-lg font-semibold bg-green-100 text-green-800 hover:bg-green-200">ساري المفعول ✅</button>
                        <button onclick="setFilterStatus('WARNING')" id="filter-btn-WARNING" class="px-3 py-1.5 rounded-lg font-semibold bg-amber-100 text-amber-800 hover:bg-amber-200">قريب الانتهاء ⚠️</button>
                        <button onclick="setFilterStatus('EXPIRED')" id="filter-btn-EXPIRED" class="px-3 py-1.5 rounded-lg font-semibold bg-red-100 text-red-800 hover:bg-red-200">منتهي الصلاحية ❌</button>
                        <button onclick="setFilterStatus('LEAVE')" id="filter-btn-LEAVE" class="px-3 py-1.5 rounded-lg font-semibold bg-cyan-100 text-cyan-800 hover:bg-cyan-200">في إجازة 🏖️</button>
                        <button onclick="toggleCategoryMenu('MALES')" id="filter-btn-MALES" class="px-3 py-1.5 rounded-lg font-semibold bg-blue-100 text-blue-800 hover:bg-blue-200">المضيفين 👨‍✈️ ▾</button>
                        <button onclick="toggleCategoryMenu('FEMALES')" id="filter-btn-FEMALES" class="px-3 py-1.5 rounded-lg font-semibold bg-pink-100 text-pink-800 hover:bg-pink-200">المضيفات 👩‍✈️ ▾</button>
                        <button onclick="toggleCategoryMenu('LEADERS')" id="filter-btn-LEADERS" class="px-3 py-1.5 rounded-lg font-semibold bg-purple-100 text-purple-800 hover:bg-purple-200">رؤساء الأطقم ⭐ ▾</button>
                    </div>

                    <!-- القوائم الفرعية للنشطين والمنتهين عند اختيار الفئة -->
                    <div id="sub-filter-container" class="hidden flex gap-2 pt-2 bg-gray-50 p-2 rounded-lg text-xs border border-gray-200">
                        <span id="sub-filter-title" class="font-bold text-gray-700 self-center">فلترة الفئة:</span>
                        <button onclick="applySubFilter('ACTIVE')" id="sub-btn-ACTIVE" class="px-3 py-1 rounded bg-green-600 text-white font-semibold">النشطين (ساري)</button>
                        <button onclick="applySubFilter('EXPIRED')" id="sub-btn-EXPIRED" class="px-3 py-1 rounded bg-red-600 text-white font-semibold">المنتهين</button>
                    </div>
                </div>

                <div id="crew-list-container" class="grid grid-cols-1 md:grid-cols-2 gap-4"></div>
            </div>

            <!-- Section: Courses Management -->
            <div id="section-courses" class="hidden space-y-4">
                <div class="bg-white p-4 rounded-xl shadow-sm border border-gray-100">
                    <h3 class="font-bold text-gray-800 mb-2">متابعة الدورات الإلزامية والكشوفات للطاقم</h3>
                    <div class="overflow-x-auto">
                        <table class="w-full text-right text-xs">
                            <thead class="bg-gray-50 text-gray-600 font-bold border-b">
                                <tr>
                                    <th class="p-2">المضيف</th>
                                    <th class="p-2">A320</th>
                                    <th class="p-2">A330</th>
                                    <th class="p-2">First Aid</th>
                                    <th class="p-2">CRM</th>
                                    <th class="p-2">الكشف الطبي</th>
                                    <th class="p-2">أمن الطيران</th>
                                    <th class="p-2">عملي A320</th>
                                    <th class="p-2">عملي A330</th>
                                    <th class="p-2">DGR</th>
                                    <th class="p-2">ID Card</th>
                                </tr>
                            </thead>
                            <tbody id="courses-table-body" class="divide-y divide-gray-100"></tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- Section: Reports & Export -->
            <div id="section-reports" class="hidden space-y-4">
                <div class="bg-white p-6 rounded-xl shadow-sm border border-gray-100 space-y-4">
                    <h3 class="font-bold text-lg text-libyan-gold">تصدير وتقارير البيانات</h3>
                    <div class="flex flex-wrap gap-3">
                        <button onclick="exportToExcel()" class="bg-green-600 hover:bg-green-700 text-white font-semibold px-5 py-2.5 rounded-lg text-sm transition flex items-center space-x-2 space-x-reverse">
                            <i class="fa-solid fa-file-excel text-lg"></i>
                            <span>تصدير كملف Excel (XLSX)</span>
                        </button>
                        <button onclick="window.print()" class="bg-gray-700 hover:bg-gray-800 text-white font-semibold px-5 py-2.5 rounded-lg text-sm transition flex items-center space-x-2 space-x-reverse">
                            <i class="fa-solid fa-print text-lg"></i>
                            <span>طباعة التقرير الحالي</span>
                        </button>
                    </div>
                </div>
            </div>
        </main>
    </div>

    <!-- Modal: Add/Edit Crew -->
    <div id="crew-modal" class="fixed inset-0 bg-black/50 hidden z-50 flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl max-w-3xl w-full max-h-[90vh] overflow-y-auto p-6 shadow-2xl space-y-4 border-t-4 border-libyan-gold">
            <div class="flex justify-between items-center border-b pb-3">
                <h3 id="modal-title" class="font-bold text-lg text-libyan-gold">إضافة مضيف جديد</h3>
                <button onclick="closeCrewModal()" class="text-gray-400 hover:text-gray-600"><i class="fa-solid fa-xmark text-xl"></i></button>
            </div>
            
            <form id="crew-form" onsubmit="saveCrew(event)" class="space-y-4 text-sm">
                <input type="hidden" id="edit-index" value="-1">
                <div class="grid grid-cols-1 md:grid-cols-3 gap-3">
                    <div>
                        <label class="block font-semibold mb-1">اسم المضيف الكامل</label>
                        <input type="text" id="form-name" required class="w-full p-2.5 border border-gray-300 rounded-lg focus:border-libyan-gold focus:outline-none">
                    </div>
                    <div>
                        <label class="block font-semibold mb-1">رقم الملف (مثال: F-101)</label>
                        <input type="text" id="form-fileNo" required class="w-full p-2.5 border border-gray-300 rounded-lg focus:border-libyan-gold focus:outline-none">
                    </div>
                    <div>
                        <label class="block font-semibold mb-1">الرقم الوظيفي</label>
                        <input type="text" id="form-empId" required class="w-full p-2.5 border border-gray-300 rounded-lg focus:border-libyan-gold focus:outline-none">
                    </div>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
                    <div>
                        <label class="block font-semibold mb-1">الطرازات المعتمدة</label>
                        <select id="form-aircraft" class="w-full p-2.5 border border-gray-300 rounded-lg focus:border-libyan-gold focus:outline-none">
                            <option value="A320 & A330">A320 & A330</option>
                            <option value="A320">A320 فقط</option>
                            <option value="A330">A330 فقط</option>
                        </select>
                    </div>
                    <div>
                        <label class="block font-semibold mb-1">ملاحظات / الوظيفة (اكتب مضيف، مضيفة، أو رئيس طاقم)</label>
                        <input type="text" id="form-notes" placeholder="مثال: مضيفة / رئيس طاقم / مضيف" class="w-full p-2.5 border border-gray-300 rounded-lg focus:border-libyan-gold focus:outline-none">
                    </div>
                </div>

                <!-- خانة الإجازة الجديدة -->
                <div class="bg-cyan-50/60 p-3 rounded-xl border border-cyan-200 flex items-center space-x-3 space-x-reverse">
                    <input type="checkbox" id="form-onLeave" class="w-5 h-5 text-cyan-600 border-gray-300 rounded focus:ring-cyan-500">
                    <label for="form-onLeave" class="font-bold text-cyan-900 cursor-pointer text-sm">في إجازة حالياً (تحديد هذه الخانة يضع المضيف في قائمة الإجازات)</label>
                </div>

                <h4 class="font-bold text-libyan-gold pt-2 border-t">تواريخ انتهاء الدورات المهنية ومهلة التنبيه (بالأيام):</h4>
                <div class="grid grid-cols-1 md:grid-cols-2 gap-3 bg-amber-50/40 p-3 rounded-xl border border-libyan-gold/20">
                    <div>
                        <label class="block text-xs font-semibold mb-1">دورة A320</label>
                        <div class="flex gap-2">
                            <input type="date" id="course-a320" class="w-full p-2 border border-gray-300 rounded-lg text-xs">
                            <input type="number" id="alert-a320" placeholder="أيام" value="60" class="w-24 p-2 border border-gray-300 rounded-lg text-xs text-center">
                        </div>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold mb-1">دورة A330</label>
                        <div class="flex gap-2">
                            <input type="date" id="course-a330" class="w-full p-2 border border-gray-300 rounded-lg text-xs">
                            <input type="number" id="alert-a330" placeholder="أيام" value="60" class="w-24 p-2 border border-gray-300 rounded-lg text-xs text-center">
                        </div>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold mb-1">الإسعافات الأولية (First Aid)</label>
                        <div class="flex gap-2">
                            <input type="date" id="course-first_aid" class="w-full p-2 border border-gray-300 rounded-lg text-xs">
                            <input type="number" id="alert-first_aid" placeholder="أيام" value="30" class="w-24 p-2 border border-gray-300 rounded-lg text-xs text-center">
                        </div>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold mb-1">إدارة موارد الطاقم (CRM)</label>
                        <div class="flex gap-2">
                            <input type="date" id="course-crm" class="w-full p-2 border border-gray-300 rounded-lg text-xs">
                            <input type="number" id="alert-crm" placeholder="أيام" value="30" class="w-24 p-2 border border-gray-300 rounded-lg text-xs text-center">
                        </div>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold mb-1">الكشف الطبي (Medical)</label>
                        <div class="flex gap-2">
                            <input type="date" id="course-medical" class="w-full p-2 border border-gray-300 rounded-lg text-xs">
                            <input type="number" id="alert-medical" placeholder="أيام" value="60" class="w-24 p-2 border border-gray-300 rounded-lg text-xs text-center">
                        </div>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold mb-1">أمن الطيران (Aviation Security)</label>
                        <div class="flex gap-2">
                            <input type="date" id="course-aviation_security" class="w-full p-2 border border-gray-300 rounded-lg text-xs">
                            <input type="number" id="alert-aviation_security" placeholder="أيام" value="30" class="w-24 p-2 border border-gray-300 rounded-lg text-xs text-center">
                        </div>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold mb-1">تدريب عملي A320</label>
                        <div class="flex gap-2">
                            <input type="date" id="course-practical_a320" class="w-full p-2 border border-gray-300 rounded-lg text-xs">
                            <input type="number" id="alert-practical_a320" placeholder="أيام" value="45" class="w-24 p-2 border border-gray-300 rounded-lg text-xs text-center">
                        </div>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold mb-1">تدريب عملي A330</label>
                        <div class="flex gap-2">
                            <input type="date" id="course-practical_a330" class="w-full p-2 border border-gray-300 rounded-lg text-xs">
                            <input type="number" id="alert-practical_a330" placeholder="أيام" value="45" class="w-24 p-2 border border-gray-300 rounded-lg text-xs text-center">
                        </div>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold mb-1">المواد الخطرة (DGR)</label>
                        <div class="flex gap-2">
                            <input type="date" id="course-dgr" class="w-full p-2 border border-gray-300 rounded-lg text-xs">
                            <input type="number" id="alert-dgr" placeholder="أيام" value="30" class="w-24 p-2 border border-gray-300 rounded-lg text-xs text-center">
                        </div>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold mb-1">بطاقة التعريف (ID Card)</label>
                        <div class="flex gap-2">
                            <input type="date" id="course-id_card" class="w-full p-2 border border-gray-300 rounded-lg text-xs">
                            <input type="number" id="alert-id_card" placeholder="أيام" value="30" class="w-24 p-2 border border-gray-300 rounded-lg text-xs text-center">
                        </div>
                    </div>
                </div>

                <div class="flex justify-end space-x-2 space-x-reverse pt-2">
                    <button type="button" onclick="closeCrewModal()" class="bg-gray-200 hover:bg-gray-300 text-gray-700 font-semibold px-4 py-2 rounded-lg text-sm transition">إلغاء</button>
                    <button type="submit" class="bg-libyan-gold hover:bg-libyan-darkGold text-white font-semibold px-6 py-2 rounded-lg text-sm transition shadow-sm">حفظ المضيف</button>
                </div>
            </form>
        </div>
    </div>

    <!-- السكريبت المحدث -->
    <script>
        const BIN_ID = "6aee1643ac6210605a73210f";
        const MASTER_KEY = "$2a$10$Th1/fnv7tUa6Sf2HOgWsHOOdTicVEJUDtVbgizqXH6C7MIs/YUcyu";
        const API_URL = `https://api.jsonbin.io/v3/b/${BIN_ID}`;

        let crewData = [];
        let currentStatusFilter = 'ALL';
        let selectedCategory = null;

        function verifyLogin() {
            const pass = document.getElementById('secret-pass').value.trim();
            if (pass === 'libyan') {
                document.getElementById('login-screen').style.display = 'none';
                document.getElementById('app-container').style.display = 'block';
                fetchDataFromCloud();
            } else {
                document.getElementById('login-error').classList.remove('hidden');
            }
        }

        function lockSystem() {
            document.getElementById('app-container').style.display = 'none';
            document.getElementById('login-screen').style.display = 'flex';
            document.getElementById('secret-pass').value = '';
        }

        async function fetchDataFromCloud() {
            try {
                const response = await fetch(API_URL, {
                    headers: { 'X-Master-Key': MASTER_KEY }
                });
                if (!response.ok) throw new Error('فشل الاتصال');
                const result = await response.json();
                crewData = result.record.crew || result.record.crwe || [];
                updateStats();
                renderCrewList();
                renderCoursesTable();
            } catch (err) {
                console.error(err);
                alert('تعذر الاتصال بقاعدة البيانات السحابية الدائمة، يرجى التحقق من الإنترنت.');
            }
        }

        async function saveChangesToCloud() {
            try {
                const response = await fetch(API_URL, {
                    method: 'PUT',
                    headers: {
                        'Content-Type': 'application/json',
                        'X-Master-Key': MASTER_KEY
                    },
                    body: JSON.stringify({ crew: crewData })
                });
                if (!response.ok) throw new Error('فشل الحفظ');
            } catch (err) {
                console.error(err);
                alert('حدث خطأ أثناء الحفظ في السحابة.');
            }
        }

        function switchTab(tabName) {
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.getElementById('tab-' + tabName).classList.add('active');
            ['crew', 'courses', 'reports'].forEach(t => {
                document.getElementById('section-' + t).classList.add('hidden');
            });
            document.getElementById('section-' + tabName).classList.remove('hidden');
        }

        function getCourseStatus(dateStr, alertDays = 30) {
            if (!dateStr) return { text: 'غير متوفر', class: 'text-gray-400' };
            const today = new Date();
            const expDate = new Date(dateStr);
            const diffTime = expDate - today;
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
            const threshold = parseInt(alertDays) || 30;

            if (diffDays < 0) {
                return { text: dateStr, class: 'text-red-600 font-bold bg-red-50 px-2 py-1 rounded' };
            } else if (diffDays <= threshold) {
                return { text: dateStr, class: 'text-amber-600 font-bold bg-amber-50 px-2 py-1 rounded' };
            } else {
                return { text: dateStr, class: 'text-green-600 font-bold bg-green-50 px-2 py-1 rounded' };
            }
        }

        function checkCrewHasExpired(c) {
            let hasExpired = false;
            const courses = c.courses || {};
            const alerts = c.alerts || {};
            Object.keys(courses).forEach(key => {
                const date = courses[key];
                if (date) {
                    const diff = Math.ceil((new Date(date) - new Date()) / (1000 * 60 * 60 * 24));
                    if (diff < 0) hasExpired = true;
                }
            });
            return hasExpired;
        }

        function updateStats() {
            let total = crewData.length;
            let males = 0, females = 0, leaders = 0, leaves = 0;
            let valid = 0, warning = 0, expired = 0;

            let malesActive = 0, malesExpired = 0;
            let femalesActive = 0, femalesExpired = 0;
            let leadersActive = 0, leadersExpired = 0;

            crewData.forEach(c => {
                const notes = (c.notes || '').toLowerCase();
                const isFemale = notes.includes('مضيفة');
                const isLeader = notes.includes('رئيس طاقم');
                const isMale = notes.includes('مضيف') && !isFemale;
                const hasExp = checkCrewHasExpired(c);

                if (isLeader) {
                    leaders++;
                    if (hasExp) leadersExpired++; else leadersActive++;
                }
                if (isFemale) {
                    females++;
                    if (hasExp) femalesExpired++; else femalesActive++;
                } else if (isMale) {
                    males++;
                    if (hasExp) malesExpired++; else malesActive++;
                }

                if (c.onLeave) {
                    leaves++;
                }

                let hasWarning = false;
                const courses = c.courses || {};
                const alerts = c.alerts || {};
                
                Object.keys(courses).forEach(key => {
                    const date = courses[key];
                    if (date) {
                        const diff = Math.ceil((new Date(date) - new Date()) / (1000 * 60 * 60 * 24));
                        const threshold = parseInt(alerts[key]) || 30;
                        if (diff <= threshold && diff >= 0) hasWarning = true;
                    }
                });

                if (hasExp) expired++;
                else if (hasWarning) warning++;
                else valid++;
            });

            document.getElementById('stat-total').innerText = total;
            
            document.getElementById('stat-males').innerText = males;
            document.getElementById('stat-males-active').innerText = malesActive;
            document.getElementById('stat-males-expired').innerText = malesExpired;

            document.getElementById('stat-females').innerText = females;
            document.getElementById('stat-females-active').innerText = femalesActive;
            document.getElementById('stat-females-expired').innerText = femalesExpired;

            document.getElementById('stat-leaders').innerText = leaders;
            document.getElementById('stat-leaders-active').innerText = leadersActive;
            document.getElementById('stat-leaders-expired').innerText = leadersExpired;

            document.getElementById('stat-leaves').innerText = leaves;
            document.getElementById('stat-valid').innerText = valid;
            document.getElementById('stat-warning').innerText = warning;
            document.getElementById('stat-expired').innerText = expired;
        }

        function setFilterStatus(status) {
            currentStatusFilter = status;
            selectedCategory = null;
            document.getElementById('sub-filter-container').classList.add('hidden');

            ['ALL', 'VALID', 'WARNING', 'EXPIRED', 'LEAVE', 'MALES', 'FEMALES', 'LEADERS'].forEach(s => {
                const btn = document.getElementById('filter-btn-' + s);
                if(btn) {
                    if(s === status) btn.className = 'px-3 py-1.5 rounded-lg font-bold bg-libyan-darkBlue text-white';
                    else btn.className = 'px-3 py-1.5 rounded-lg font-semibold bg-gray-200 text-gray-700 hover:bg-gray-300';
                }
            });
            renderCrewList();
        }

        function toggleCategoryMenu(category) {
            selectedCategory = category;
            currentStatusFilter = category;
            document.getElementById('sub-filter-container').classList.remove('hidden');
            
            ['ALL', 'VALID', 'WARNING', 'EXPIRED', 'LEAVE', 'MALES', 'FEMALES', 'LEADERS'].forEach(s => {
                const btn = document.getElementById('filter-btn-' + s);
                if(btn) {
                    if(s === category) btn.className = 'px-3 py-1.5 rounded-lg font-bold bg-libyan-darkBlue text-white';
                    else btn.className = 'px-3 py-1.5 rounded-lg font-semibold bg-gray-200 text-gray-700 hover:bg-gray-300';
                }
            });
            
            const titles = { 'MALES': 'المضيفين', 'FEMALES': 'المضيفات', 'LEADERS': 'رؤساء الأطقم' };
            document.getElementById('sub-filter-title').innerText = `فلترة ${titles[category]}:`;
            applySubFilter('ACTIVE');
        }

        function applySubFilter(subType) {
            ['ACTIVE', 'EXPIRED'].forEach(s => {
                const btn = document.getElementById('sub-btn-' + s);
                if(s === subType) btn.className = 'px-3 py-1 rounded bg-libyan-gold text-white font-bold';
                else btn.className = 'px-3 py-1 rounded bg-gray-300 text-gray-700 font-semibold';
            });
            currentStatusFilter = selectedCategory + '_' + subType;
            renderCrewList();
        }

        function renderCrewList() {
            const container = document.getElementById('crew-list-container');
            const search = document.getElementById('search-input').value.toLowerCase();
            const aircraftFilter = document.getElementById('filter-aircraft').value;

            const filtered = crewData.filter(c => {
                const matchSearch = (c.name && c.name.toLowerCase().includes(search)) || (c.fileNo && c.fileNo.toLowerCase().includes(search)) || (c.empId && c.empId.toLowerCase().includes(search));
                const matchAircraft = aircraftFilter === 'ALL' || (c.aircraft && c.aircraft.includes(aircraftFilter));
                
                if (!matchSearch || !matchAircraft) return false;

                const notes = (c.notes || '').toLowerCase();
                const isFemale = notes.includes('مضيفة');
                const isLeader = notes.includes('رئيس طاقم');
                const isMale = notes.includes('مضيف') && !isFemale;
                const hasExp = checkCrewHasExpired(c);

                if (currentStatusFilter === 'VALID') {
                    if (hasExp) return false;
                } else if (currentStatusFilter === 'WARNING') {
                    let hasWarn = false;
                    Object.keys(c.courses || {}).forEach(k => {
                        const d = c.courses[k];
                        if (d) {
                            const diff = Math.ceil((new Date(d) - new Date()) / (1000 * 60 * 60 * 24));
                            const th = parseInt(c.alerts?.[k]) || 30;
                            if (diff <= th && diff >= 0) hasWarn = true;
                        }
                    });
                    if (!hasWarn || hasExp) return false;
                } else if (currentStatusFilter === 'EXPIRED') {
                    if (!hasExp) return false;
                } else if (currentStatusFilter === 'LEAVE') {
                    if (!c.onLeave) return false;
                } else if (currentStatusFilter === 'MALES_ACTIVE') {
                    if (!isMale || hasExp) return false;
                } else if (currentStatusFilter === 'MALES_EXPIRED') {
                    if (!isMale || !hasExp) return false;
                } else if (currentStatusFilter === 'FEMALES_ACTIVE') {
                    if (!isFemale || hasExp) return false;
                } else if (currentStatusFilter === 'FEMALES_EXPIRED') {
                    if (!isFemale || !hasExp) return false;
                } else if (currentStatusFilter === 'LEADERS_ACTIVE') {
                    if (!isLeader || hasExp) return false;
                } else if (currentStatusFilter === 'LEADERS_EXPIRED') {
                    if (!isLeader || !hasExp) return false;
                }
                return true;
            });

            if (filtered.length === 0) {
                container.innerHTML = '<div class="col-span-2 text-center py-10 text-gray-500 bg-white rounded-xl shadow-sm">لا توجد بيانات مطابقة للعرض</div>';
                return;
            }

            container.innerHTML = filtered.map((c) => {
                const globalIndex = crewData.indexOf(c);
                
                let hasExpired = false;
                let hasWarning = false;
                const courses = c.courses || {};
                const alerts = c.alerts || {};

                Object.keys(courses).forEach(key => {
                    const date = courses[key];
                    if (date) {
                        const diff = Math.ceil((new Date(date) - new Date()) / (1000 * 60 * 60 * 24));
                        const threshold = parseInt(alerts[key]) || 30;
                        if (diff < 0) hasExpired = true;
                        else if (diff <= threshold) hasWarning = true;
                    }
                });

                let dotColor = 'bg-green-500';
                let statusTitle = 'ساري المفعول';
                if (hasExpired) {
                    dotColor = 'bg-red-500';
                    statusTitle = 'يوجد دورات منتهية';
                } else if (hasWarning) {
                    dotColor = 'bg-amber-500';
                    statusTitle = 'يوجد دورات قريبة الانتهاء';
                }

                return `
                    <div class="bg-white p-4 rounded-xl shadow-sm border border-gray-100 flex justify-between items-center">
                        <div>
                            <div class="flex items-center space-x-2 space-x-reverse">
                                <span class="w-3 h-3 rounded-full ${dotColor} inline-block shrink-0 shadow-sm" title="${statusTitle}"></span>
                                <h4 class="font-bold text-gray-800 text-base">${c.name || 'بدون اسم'}</h4>
                                ${c.onLeave ? '<span class="bg-cyan-100 text-cyan-800 text-[10px] px-2 py-0.5 rounded-full font-bold">في إجازة 🏖️</span>' : ''}
                            </div>
                            <p class="text-xs text-gray-500 mt-1">الملف: ${c.fileNo || '-'} | الوظيفي: ${c.empId || '-'} | الطراز: ${c.aircraft || '-'}</p>
                            <p class="text-xs text-libyan-gold mt-0.5">${c.notes || ''}</p>
                        </div>
                        <div class="text-left space-y-2">
                            <div class="flex gap-1 justify-end pt-1">
                                <button onclick="openEditCrewModal(${globalIndex})" class="p-2 text-blue-600 hover:bg-blue-50 rounded-lg transition" title="تعديل الملف"><i class="fa-solid fa-pen-to-square text-sm"></i></button>
                                <button onclick="deleteCrew(${globalIndex})" class="p-2 text-red-600 hover:bg-red-50 rounded-lg transition" title="حذف"><i class="fa-solid fa-trash text-sm"></i></button>
                            </div>
                        </div>
                    </div>
                `;
            }).join('');
        }

        function renderCoursesTable() {
            const tbody = document.getElementById('courses-table-body');
            if (crewData.length === 0) {
                tbody.innerHTML = '<tr><td colspan="11" class="text-center py-6 text-gray-400">لا توجد سجلات مضافة حتى الآن</td></tr>';
                return;
            }
            tbody.innerHTML = crewData.map(c => {
                const co = c.courses || {};
                const al = c.alerts || {};
                return `
                    <tr class="hover:bg-gray-50 text-center">
                        <td class="p-2 font-bold text-right">${c.name || '-'} ${c.onLeave ? '<span class="text-cyan-600 text-[10px]">(إجازة)</span>' : ''}</td>
                        <td class="p-2"><span class="${getCourseStatus(co.a320, al.a320).class}">${getCourseStatus(co.a320, al.a320).text}</span></td>
                        <td class="p-2"><span class="${getCourseStatus(co.a330, al.a330).class}">${getCourseStatus(co.a330, al.a330).text}</span></td>
                        <td class="p-2"><span class="${getCourseStatus(co.first_aid, al.first_aid).class}">${getCourseStatus(co.first_aid, al.first_aid).text}</span></td>
                        <td class="p-2"><span class="${getCourseStatus(co.crm, al.crm).class}">${getCourseStatus(co.crm, al.crm).text}</span></td>
                        <td class="p-2"><span class="${getCourseStatus(co.medical, al.medical).class}">${getCourseStatus(co.medical, al.medical).text}</span></td>
                        <td class="p-2"><span class="${getCourseStatus(co.aviation_security, al.aviation_security).class}">${getCourseStatus(co.aviation_security, al.aviation_security).text}</span></td>
                        <td class="p-2"><span class="${getCourseStatus(co.practical_a320, al.practical_a320).class}">${getCourseStatus(co.practical_a320, al.practical_a320).text}</span></td>
                        <td class="p-2"><span class="${getCourseStatus(co.practical_a330, al.practical_a330).class}">${getCourseStatus(co.practical_a330, al.practical_a330).text}</span></td>
                        <td class="p-2"><span class="${getCourseStatus(co.dgr, al.dgr).class}">${getCourseStatus(co.dgr, al.dgr).text}</span></td>
                        <td class="p-2"><span class="${getCourseStatus(co.id_card, al.id_card).class}">${getCourseStatus(co.id_card, al.id_card).text}</span></td>
                    </tr>
                `;
            }).join('');
        }

        function openAddCrewModal() {
            document.getElementById('edit-index').value = "-1";
            document.getElementById('modal-title').innerText = "إضافة مضيف جديد";
            document.getElementById('crew-form').reset();
            document.getElementById('form-onLeave').checked = false;
            document.getElementById('crew-modal').classList.remove('hidden');
        }

        function openEditCrewModal(index) {
            const c = crewData[index];
            document.getElementById('edit-index').value = index;
            document.getElementById('modal-title').innerText = "تعديل بيانات المضيف";
            
            document.getElementById('form-name').value = c.name || '';
            document.getElementById('form-fileNo').value = c.fileNo || '';
            document.getElementById('form-empId').value = c.empId || '';
            document.getElementById('form-aircraft').value = c.aircraft || 'A320 & A330';
            document.getElementById('form-notes').value = c.notes || '';
            document.getElementById('form-onLeave').checked = !!c.onLeave;

            const co = c.courses || {};
            const al = c.alerts || {};

            const keys = ['a320', 'a330', 'first_aid', 'crm', 'medical', 'aviation_security', 'practical_a320', 'practical_a330', 'dgr', 'id_card'];
            keys.forEach(k => {
                document.getElementById('course-' + k).value = co[k] || '';
                if(document.getElementById('alert-' + k)) {
                    document.getElementById('alert-' + k).value = al[k] || (k.includes('a320') || k.includes('a330') || k.includes('medical') ? 60 : 30);
                }
            });

            document.getElementById('crew-modal').classList.remove('hidden');
        }

        function closeCrewModal() {
            document.getElementById('crew-modal').classList.add('hidden');
        }

        async function saveCrew(event) {
            event.preventDefault();
            const editIndex = parseInt(document.getElementById('edit-index').value);

            const crewObj = {
                name: document.getElementById('form-name').value,
                fileNo: document.getElementById('form-fileNo').value,
                empId: document.getElementById('form-empId').value,
                aircraft: document.getElementById('form-aircraft').value,
                notes: document.getElementById('form-notes').value,
                onLeave: document.getElementById('form-onLeave').checked,
                courses: {
                    a320: document.getElementById('course-a320').value,
                    a330: document.getElementById('course-a330').value,
                    first_aid: document.getElementById('course-first_aid').value,
                    crm: document.getElementById('course-crm').value,
                    medical: document.getElementById('course-medical').value,
                    aviation_security: document.getElementById('course-aviation_security').value,
                    practical_a320: document.getElementById('course-practical_a320').value,
                    practical_a330: document.getElementById('course-practical_a330').value,
                    dgr: document.getElementById('course-dgr').value,
                    id_card: document.getElementById('course-id_card').value
                },
                alerts: {
                    a320: document.getElementById('alert-a320').value,
                    a330: document.getElementById('alert-a330').value,
                    first_aid: document.getElementById('alert-first_aid').value,
                    crm: document.getElementById('alert-crm').value,
                    medical: document.getElementById('alert-medical').value,
                    aviation_security: document.getElementById('alert-aviation_security').value,
                    practical_a320: document.getElementById('alert-practical_a320').value,
                    practical_a330: document.getElementById('alert-practical_a330').value,
                    dgr: document.getElementById('alert-dgr').value,
                    id_card: document.getElementById('alert-id_card').value
                }
            };

            if (editIndex === -1) {
                crewData.push(crewObj);
            } else {
                crewData[editIndex] = crewObj;
            }

            closeCrewModal();
            updateStats();
            renderCrewList();
            renderCoursesTable();
            await saveChangesToCloud();
        }

        async function deleteCrew(index) {
            if (confirm('هل أنت متأكد من حذف هذا المضيف نهائياً؟')) {
                crewData.splice(index, 1);
                updateStats();
                renderCrewList();
                renderCoursesTable();
                await saveChangesToCloud();
            }
        }

        function exportToExcel() {
            const flatData = crewData.map(c => ({
                'اسم المضيف': c.name,
                'رقم الملف': c.fileNo,
                'الرقم الوظيفي': c.empId,
                'الطرازات': c.aircraft,
                'ملاحظات': c.notes,
                'الحالة': c.onLeave ? 'في إجازة' : 'على رأس العمل',
                'A320': c.courses?.a320 || '',
                'A330': c.courses?.a330 || '',
                'First Aid': c.courses?.first_aid || '',
                'CRM': c.courses?.crm || '',
                'Medical': c.courses?.medical || '',
                'Aviation Security': c.courses?.aviation_security || '',
                'Practical A320': c.courses?.practical_a320 || '',
                'Practical A330': c.courses?.practical_a330 || '',
                'DGR': c.courses?.dgr || '',
                'ID Card': c.courses?.id_card || ''
            }));
            const ws = XLSX.utils.json_to_sheet(flatData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "CrewRecords");
            XLSX.writeFile(wb, "Libyan_Crew_Management_System.xlsx");
        }
    </script>
</body>
</html>
