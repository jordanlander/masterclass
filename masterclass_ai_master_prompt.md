# 🧠 Master Presentation Prompt Template
Copy the entire block below into any AI chat, then fill in the placeholders.

---

title: "[Presentation Title]"
presenter: "[Your Name]"
org: "[Organization/School]"
date: "[Date]"

---

## ✅ Instructions for the AI

1. **Output Format**
   - Deliver the ENTIRE presentation inside **one** markdown code block.
   - Use a YAML-style header (as above).
   - Each slide is an H2 or H3 heading (`##` / `###`).
   - Do **not** label with “Slide 1, Slide 2.”
   - If numbering helps you, put it in a `NOTE:` line; it must not appear on slides.

2. **Content Expectations**
   - Target: **[Number]** slides (minimum).
   - Sections to cover (10 typical sections, modify as needed):
     1. Introduction
     2. Context & Data
     3. [Key Principle 1]
     4. Brain Science / Theory
     5. Relationships / Culture
     6. Classroom Strategies
     7. SEL / Practice
     8. Family & Community
     9. Educator Wellness
     10. Closing & Resources
   - **Every bullet must be a full sentence** or short paragraph—no one-word bullets.
   - Each slide should read like a mini teaching point and stand on its own.
   - Include reflective questions, quotes, scenarios, or role-play prompts to boost engagement.

3. **Media Integration**
   - Embed media directly in slides (not at the end):
     - **Images:** `![](https://picsum.photos/seed/[topic]/1000/450)` or real URLs provided.
     - **Videos:** `VIDEO: [Title](https://www.youtube.com/watch?v=XXXXXX)`
     - **Tables:** Inline markdown tables with clear labels.
   - Mix text, images, video links, tables, quotes, and questions to avoid repetition.

4. **Research & Credibility**
   - Pull from trusted, current sources (CDC, CASEL, SAMHSA, etc.).
   - Explain *why* each fact matters to the audience.
   - Highlight location-specific data if relevant.

5. **Tone & Style**
   - Use professional yet approachable language.
   - Avoid jargon unless defined.
   - Use emojis sparingly to create emotional engagement (🌟, 💭, etc.).
   - Insert pull quotes from experts to break up text.

6. **Flow**
   - Each section should progress logically:
     - definition → data → quote → strategy → reflection → example.
   - Include interactive elements (practice scenarios, role-plays, “What would you do?” questions).

7. **Finishing Touches**
   - End with key takeaways, quotes, and resource links.
   - Check spelling, grammar, and consistent formatting before finalizing.

## 📋 Command List for Media

Use these commands within slide content to embed multimedia:

| To insert | Syntax | Example |
|-----------|--------|---------|
| Image or picture | `![](IMAGE_URL)` | `![](https://picsum.photos/seed/topic/1000/450)` |
| Image with alt text | `![Alt Text](IMAGE_URL)` | `![Brain Diagram](https://example.com/brain.png)` |
| Video | `VIDEO: [Title](VIDEO_URL)` | `VIDEO: Growth Mindset Explained(https://youtu.be/abcdef)` |

## 🔧 Prompt to the AI

Using the format and instructions above:

"Create a [target number]-slide masterclass on **[Topic]** for **[Audience]**.
Follow the YAML header and slide structure.
Include full sentences, current research, quotes, reflective questions, and real or placeholder media links.
Target sections: [list your 10 sections or modify].
Deliver everything in one markdown code block."

