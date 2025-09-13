# MailMute   

**AI-Powered Email Content Analyzer & Smart Unsubscriber**  

MailMute is a desktop application designed to help users **analyze emails, detect unsubscribe opportunities, and manage mailing lists intelligently**. Using AI-inspired pattern matching, it identifies unsubscribe links, evaluates sender credibility, and automates or guides the unsubscription process.  

---

##  Features  

- **AI-Powered Analysis**: Detects unsubscribe links and headers in emails using pattern matching and heuristics.  
- **Smart Unsubscribe**: Attempts unsubscription via email or web links safely.  
- **Database Management**: Tracks unsubscribe history and sender statistics using SQLite.  
- **Confidence Scoring**: Assigns a confidence score to each unsubscribe candidate for better reliability.  
- **Modern GUI**: Built with Tkinter for a scrollable, modern interface.  

---

##  Tech Stack  

- **Python 3.10+**  
- **Tkinter** for GUI  
- **BeautifulSoup4** for HTML parsing  
- **SQLite** for storing email history  
- **Requests** for web interactions  

---

## 🖼 Screenshots  

<table>
  <tr>
    <td><img width="326" height="1168" src="https://github.com/user-attachments/assets/a36855f4-5064-4736-94c5-ca2250271560" /></td>
    <td><img width="326" height="1168" src="https://github.com/user-attachments/assets/abe5cd8a-b4c7-492a-b2cf-9c1ec0e370af" /></td>
  </tr>
  <tr>
    <td><img width="326" height="1168" src="https://github.com/user-attachments/assets/6a435b4e-2825-4c83-a414-b65646e822da" /></td>
    <td><img width="326" height="1168" src="https://github.com/user-attachments/assets/8c1679a4-7363-4b04-ad04-10b446a0937b" /></td>
  </tr>
</table>

---

## How It Works  

1. **Analyze Email Content**  
   - Extracts HTML and text content from emails.  
   - Identifies unsubscribe links and headers.  
   - Calculates a confidence score for each detection.  

2. **Unsubscribe Automation**  
   - Attempts to unsubscribe via email headers or web links.  
   - Records actions and results in a local SQLite database.  
   - Tracks sender statistics for better decision-making over time.  

3. **Modern Scrollable GUI**  
   - Intuitive interface to browse email history.  
   - Scrollable frames and clean visual design using Tkinter.  

---

## ⚙ Usage  
Generate your Mail-Password on Google Account
Clone the repository:  
```bash
git clone https://github.com/yourusername/MailMute.git
cd MailMute''
