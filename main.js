document.addEventListener('DOMContentLoaded', () => {
    const ballContainer = document.getElementById('ball-container');
    const generateBtn = document.getElementById('generate-btn');
    const themeToggle = document.getElementById('theme-toggle');
    const themeIcon = themeToggle.querySelector('.icon');

    // Theme Logic
    const currentTheme = localStorage.getItem('theme') || 'dark';
    if (currentTheme === 'light') {
        document.body.classList.add('light-mode');
        themeIcon.textContent = '☀️';
    }

    themeToggle.addEventListener('click', () => {
        document.body.classList.toggle('light-mode');
        const isLight = document.body.classList.contains('light-mode');
        localStorage.setItem('theme', isLight ? 'light' : 'dark');
        themeIcon.textContent = isLight ? '☀️' : '🌙';
    });

    /**
     * Generates 6 unique random numbers between 1 and 45.
     * @returns {number[]} Sorted array of 6 numbers.
     */
    function generateLottoNumbers() {
        const numbers = new Set();
        while (numbers.size < 6) {
            const randomNum = Math.floor(Math.random() * 45) + 1;
            numbers.add(randomNum);
        }
        return Array.from(numbers).sort((a, b) => a - b);
    }

    /**
     * Determines the CSS class for a ball based on its number range.
     * @param {number} num 
     * @returns {string} CSS class name
     */
    function getBallClass(num) {
        if (num <= 10) return 'ball-1';
        if (num <= 20) return 'ball-2';
        if (num <= 30) return 'ball-3';
        if (num <= 40) return 'ball-4';
        return 'ball-5';
    }

    /**
     * Clears the container and renders new balls with animation.
     */
    async function renderBalls() {
        // Disable button during animation
        generateBtn.disabled = true;
        
        // Clear previous balls
        ballContainer.innerHTML = '';
        
        const numbers = generateLottoNumbers();
        
        for (let i = 0; i < numbers.length; i++) {
            const num = numbers[i];
            
            // Create ball element
            const ball = document.createElement('div');
            ball.className = `lotto-ball ${getBallClass(num)} pop`;
            ball.textContent = num;
            
            // Add to container
            ballContainer.appendChild(ball);
            
            // Wait for a short delay to create sequential effect
            await new Promise(resolve => setTimeout(resolve, 300));
        }
        
        // Re-enable button
        generateBtn.disabled = false;
    }

    // Event Listener
    generateBtn.addEventListener('click', renderBalls);
});
