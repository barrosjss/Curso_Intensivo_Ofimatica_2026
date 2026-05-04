import { courseData } from './data.js';

let currentDay = 'home';

const sidebarNav = document.getElementById('sidebar-nav');
const courseContent = document.getElementById('course-content');

function init() {
    // 1. Try to get day from URL hash (e.g., #day=1)
    const hash = window.location.hash;
    if (hash && hash.includes('day=')) {
        currentDay = parseInt(hash.split('=')[1]);
    } 
    // 2. If no hash, try to get from localStorage
    else {
        const savedDay = localStorage.getItem('lastCourseDay');
        if (savedDay) {
            currentDay = savedDay === 'home' ? 'home' : parseInt(savedDay);
        }
    }
    
    setDay(currentDay);

    // Listen for hash changes (back/forward buttons)
    window.addEventListener('hashchange', () => {
        const newHash = window.location.hash;
        if (newHash.includes('day=')) {
            const day = parseInt(newHash.split('=')[1]);
            if (day !== currentDay) setDay(day);
        } else if (newHash === '#home' || newHash === '') {
            if (currentDay !== 'home') setDay('home');
        }
    });
}

function renderSidebar() {
    sidebarNav.innerHTML = `
        <div class="nav-item ${currentDay === 'home' ? 'active' : ''}" onclick="window.setDay('home')">
            <span class="day-badge">INICIO</span>
            <span>General</span>
        </div>
    ` + courseData.map(day => `
        <div class="nav-item ${day.day === currentDay ? 'active' : ''}" onclick="window.setDay(${day.day})">
            <span class="day-badge">DÍA ${day.day}</span>
            <span>${day.tool}</span>
        </div>
    `).join('');
}

window.setDay = function(day) {
    currentDay = day;
    
    // Save state
    localStorage.setItem('lastCourseDay', day);
    window.location.hash = day === 'home' ? 'home' : `day=${day}`;
    
    renderSidebar();
    if (day === 'home') {
        renderHome();
        updateProgress(0);
    } else {
        renderDay(currentDay);
        updateProgress((day / courseData.length) * 100);
    }
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

function updateProgress(percent) {
    const bar = document.getElementById('progress-bar');
    if (bar) bar.style.width = percent + '%';
}

function renderHome() {
    import('./data.js').then(({ evaluationCriteria }) => {
        courseContent.innerHTML = `
            <div class="hero-section">
                <img src="assets/hero.png" alt="Hero Office">
                <div class="hero-overlay">
                    <div class="day-info">CURSO INTENSIVO 2026</div>
                    <h2 style="font-size: 2.5rem; margin-top: 0.5rem;">Domina el Ecosistema Microsoft 365</h2>
                </div>
            </div>

            <div class="section-card">
                <h3 class="section-title">Sobre el Curso</h3>
                <p style="line-height: 1.6; color: var(--text-secondary); font-size: 1.1rem;">
                    Este curso intensivo está diseñado para desarrollar competencias prácticas en las principales herramientas de productividad, 
                    orientadas a optimizar los procesos administrativos, de comunicación y análisis de datos en organizaciones.
                </p>
            </div>

            <div class="section-card">
                <h3 class="section-title">Criterios de Evaluación</h3>
                <div style="overflow-x: auto;">
                    <table style="width: 100%; border-collapse: collapse; margin-top: 1rem;">
                        <thead>
                            <tr style="border-bottom: 2px solid var(--glass-border); text-align: left;">
                                <th style="padding: 1rem; color: var(--accent-blue);">Componente</th>
                                <th style="padding: 1rem; color: var(--accent-blue);">Peso</th>
                                <th style="padding: 1rem; color: var(--accent-blue);">Descripción</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${evaluationCriteria.map(item => `
                                <tr style="border-bottom: 1px solid var(--glass-border);">
                                    <td style="padding: 1rem; font-weight: 600;">${item.component}</td>
                                    <td style="padding: 1rem;"><span style="background: var(--accent-blue); padding: 0.2rem 0.6rem; border-radius: 1rem; font-size: 0.8rem;">${item.weight}</span></td>
                                    <td style="padding: 1rem; color: var(--text-secondary); font-size: 0.9rem;">${item.description}</td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>
            </div>

            <div style="text-align: center; margin-top: 2rem; display: flex; justify-content: center; gap: 1rem;">
                <button class="btn btn-primary" style="padding: 1rem 3rem;" onclick="setDay(1)">
                    Empezar Curso →
                </button>
                <a href="recursos/Curso_Intensivo_Ofimatica_2026.docx.pdf" download class="btn btn-secondary" style="padding: 1rem 2rem; text-decoration: none;">
                    📂 Descargar PDF
                </a>
            </div>
        `;
    });
}

function renderDay(dayId) {
    const data = courseData.find(d => d.day === dayId);
    
    courseContent.innerHTML = `
        <div class="hero-section">
            <img src="assets/hero.png" alt="Hero Office">
            <div class="hero-overlay">
                <div class="day-info">DÍA ${data.day} • ${data.tool}</div>
                <h2 style="font-size: 2.5rem; margin-top: 0.5rem;">${data.title}</h2>
            </div>
        </div>

        <div class="section-card">
            <h3 class="section-title">Objetivo de la Sesión</h3>
            <p style="line-height: 1.6; color: var(--text-secondary); font-size: 1.1rem;">${data.objective}</p>
        </div>

        <div class="section-card">
            <h3 class="section-title">Lecciones del Día</h3>
            ${data.lessons.map(lesson => `
                <div class="lesson-item" style="margin-bottom: 2.5rem; border-left: 2px solid var(--glass-border); padding-left: 1.5rem;">
                    <h4 style="color: var(--accent-blue); margin-bottom: 0.75rem; font-size: 1.2rem;">${lesson.title}</h4>
                    <p style="line-height: 1.7; color: var(--text-primary); margin-bottom: 1rem;">${lesson.content}</p>
                    
                    ${lesson.videoUrl ? `
                        <div style="margin: 1.5rem 0; border-radius: 0.75rem; overflow: hidden; box-shadow: 0 4px 20px rgba(0,0,0,0.4); max-width: ${lesson.videoUrl.includes('tiktok') ? '325px' : '100%'}; margin-left: auto; margin-right: auto;">
                            <iframe width="100%" height="${lesson.videoUrl.includes('tiktok') ? '580' : '315'}" src="${lesson.videoUrl}" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe>
                        </div>
                    ` : ''}

                    ${lesson.videoUrls ? lesson.videoUrls.map(url => `
                        <div style="margin: 1.5rem 0; border-radius: 0.75rem; overflow: hidden; box-shadow: 0 4px 20px rgba(0,0,0,0.4); max-width: ${url.includes('tiktok') ? '325px' : '100%'}; margin-left: auto; margin-right: auto;">
                            <iframe width="100%" height="${url.includes('tiktok') ? '580' : '315'}" src="${url}" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe>
                        </div>
                    `).join('') : ''}

                    ${lesson.imageUrl ? `
                        <div style="margin: 1.5rem 0; border-radius: 0.75rem; overflow: hidden; box-shadow: 0 4px 20px rgba(0,0,0,0.4);">
                            <img src="${lesson.imageUrl}" alt="${lesson.title}" style="width: 100%; display: block;">
                        </div>
                    ` : ''}

                    <div style="background: rgba(168, 85, 247, 0.1); border: 1px solid rgba(168, 85, 247, 0.2); padding: 0.75rem 1rem; border-radius: 0.5rem; font-size: 0.9rem;">
                        <span style="color: var(--accent-purple); font-weight: bold;">💡 Tip del Instructor:</span> 
                        <span style="color: var(--text-secondary);">${lesson.tip}</span>
                    </div>
                </div>
            `).join('')}
        </div>

        <div class="section-card">
            <h3 class="section-title">Actividad Práctica</h3>
            <p style="line-height: 1.6; color: var(--text-secondary);">${data.practice}</p>
        </div>

        <div class="section-card quiz-container">
            <h3 class="section-title">Pon a prueba tu conocimiento</h3>
            <p style="margin-bottom: 1.5rem;">${data.quiz.question}</p>
            <div id="quiz-options">
                ${data.quiz.options.map((option, index) => `
                    <button class="quiz-option" onclick="checkAnswer(${index}, ${data.quiz.correct})">
                        ${option}
                    </button>
                `).join('')}
            </div>
            <div id="quiz-feedback" style="margin-top: 1rem; font-weight: 600; display: none;"></div>
        </div>

        <div class="step-nav">
            <button class="btn btn-secondary" onclick="setDay(${dayId > 1 ? dayId - 1 : dayId})" ${dayId === 1 ? 'style="visibility: hidden;"' : ''}>
                ← Anterior
            </button>
            <button class="btn btn-primary" onclick="setDay(${dayId < 10 ? dayId + 1 : dayId})" ${dayId === 10 ? 'style="visibility: hidden;"' : ''}>
                Siguiente Sesión →
            </button>
        </div>
    `;
}

window.checkAnswer = function(selectedIndex, correctIndex) {
    const buttons = document.querySelectorAll('.quiz-option');
    const feedback = document.getElementById('quiz-feedback');
    
    buttons.forEach((btn, idx) => {
        btn.disabled = true;
        if (idx === correctIndex) {
            btn.classList.add('correct');
        } else if (idx === selectedIndex) {
            btn.classList.add('wrong');
        }
    });

    feedback.style.display = 'block';
    if (selectedIndex === correctIndex) {
        feedback.innerHTML = '¡Correcto! Excelente trabajo.';
        feedback.style.color = 'var(--accent-green)';
    } else {
        feedback.innerHTML = 'Incorrecto. La respuesta correcta está resaltada en verde.';
        feedback.style.color = '#ef4444';
    }
}

init();
