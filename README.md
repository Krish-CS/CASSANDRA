# 🌸 Cassandra - AI Presentation Generator

Cassandra is an intelligent, Flask-based web application that automatically generates beautiful, fully-structured PowerPoint (.pptx) presentations from a single topic prompt. It leverages AI for high-quality content generation and the Pexels API for stunning, customizable background imagery.

## ✨ Features

- **Topic-to-PPT in Seconds (Flash Mode):** Simply enter a topic, and Cassandra will generate a complete 10-20 slide presentation and download it instantly.
- **Interactive Editor (Decide Mode):** Preview your slides, edit titles, rewrite content line-by-line, and change background themes before downloading the final file.
- **Dynamic Backgrounds:** Integrates with the Pexels API to pull high-quality, color-themed background images for your slides.
- **Smart Formatting:** AI automatically determines whether a slide should contain a paragraph (for introductions/conclusions) or bullet points.
- **Mobile Responsive:** A sleek, dark-themed UI that works flawlessly on desktop and mobile devices.

## 🚀 Technology Stack

- **Backend:** Python 3, Flask, Gunicorn
- **PPT Generation:** `python-pptx`
- **AI Content:** Groq API (or compatible LLM endpoints)
- **Imagery:** Pexels API
- **Frontend:** HTML5, CSS3, Vanilla JavaScript

## 💻 Local Execution Environment Setup

### 1. Clone the repository
```bash
git clone https://github.com/Krish-CS/CASSANDRA.git
cd CASSANDRA
```

### 2. Create a Virtual Environment (Recommended)
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\\Scripts\\activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

### 4. Setup Environment Variables
Create a file named `.env.ppt` in the root directory and add your API keys:
```env
GROQ_API_KEY=your_groq_api_key_here
PEXELS_API_KEY=your_pexels_api_key_here
```
*(Note: If you are using a different LLM provider, you may need to adjust the environment variables according to `presentation/slide_generator.py`)*

### 5. Run the Application
```bash
python app.py
```
Open your browser and navigate to `http://localhost:5000`.

## 🌍 Production Deployment (Render)

Cassandra is optimized for free-tier deployment on platforms like [Render](https://render.com). 

1. Create a new **Web Service** on Render and connect this GitHub repository.
2. Set the Environment to **Python 3**.
3. **Build Command:** `pip install -r requirements.txt`
4. **Start Command:** `gunicorn app:app`
5. **Environment Variables:** Add your `GROQ_API_KEY` and `PEXELS_API_KEY` in the Render dashboard.
6. Click **Deploy**.

> **Note on Free Tiers:** Cassandra includes a `/ping` endpoint. You can use a free service like [UptimeRobot](https://uptimerobot.com) to ping `https://your-app-url.onrender.com/ping` every 5-10 minutes to prevent the Render instance from spinning down due to inactivity.

## 🧹 Storage Management

To prevent server storage from filling up during production hosting, Cassandra includes:
- **Immediate Cleanup:** PPTs generated in Flash mode and downloaded are immediately deleted from the server.
- **Background Scheduler:** A background thread routinely checks the `output/` folder and securely deletes any temporary `.pptx` files older than 30 minutes.

## 📝 License

This project is licensed under the MIT License.
