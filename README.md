📇 VFC2Excel

Seamlessly convert VCF (vCard) files into Excel spreadsheets — smart, slick, and speedy!

VFC2Excel is a handy utility that transforms contact data from .vcf (vCard) format into clean, structured Excel sheets. Perfect for managing phone contacts, backups, or bulk imports.

✨ Features

📂 VCF-to-Excel Conversion – Convert .vcf files into .xlsx in seconds.

📦 Batch Processing – Handle multiple VCF files at once.

📝 Preserves Data – Names, numbers, emails, birthdays, and more.

⚡ Lightweight & Fast – Built with efficiency in mind.

🛠 Customizable Mapping – Optional fields supported (addresses, notes, etc.).

🛠 Tech Stack

Language: Python 🐍

Libraries: pandas, openpyxl, vobject 

🚀 Getting Started

1️⃣ Clone the repo

git clone https://github.com/Charanvas/VFC2Excel.git
cd VFC2Excel


2️⃣ Install dependencies

pip install -r requirements.txt


3️⃣ Run the tool

python vfc2excel.py --input contacts.vcf --output contacts.xlsx


4️⃣ Batch convert folder

python vfc2excel.py --input-dir ./vcfs --output-dir ./excels

📸 Example
python vfc2excel.py -i sample.vcf -o sample.xlsx


Output: sample.xlsx containing all contacts in a clean tabular format.

🤝 Contributing

We welcome ideas & improvements:

Fork 🍴 the repo

Create a branch 🌱 (git checkout -b feature/my-feature)

Commit 📝 (git commit -m "Add batch conversion")

Push 🚀 (git push origin feature/my-feature)

Open a PR 🎉
