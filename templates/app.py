from flask import Flask, render_template, request
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

@app.route("/")
def top():
    return render_template("top.html")


@app.route("/diagnosis")
def diagnosis():
    return render_template("index.html")


@app.route("/result", methods=["POST"])
def result():

    company = request.form["company"]

    contact = 0
    recruit = 0
    recognition = 0
    organization = 0

    answers = {}

    # 接点
    for i in range(1,7):
        val = int(request.form.get(f"contact{i}",0))
        contact += val
        answers[f"contact{i}"] = val

    # 求人力
    for i in range(1,7):
        val = int(request.form.get(f"recruit{i}",0))
        recruit += val
        answers[f"recruit{i}"] = val

    # 認知
    for i in range(1,7):
        val = int(request.form.get(f"recognition{i}",0))
        recognition += val
        answers[f"recognition{i}"] = val

    # 組織力
    for i in range(1,7):
        val = int(request.form.get(f"organization{i}",0))
        organization += val
        answers[f"organization{i}"] = val


    scores = {
        "接点不足型": contact,
        "求人改善型": recruit,
        "認知不足型": recognition,
        "組織改善型": organization
    }

    result_type = min(scores, key=scores.get)


    comments = {
        "接点不足型":"求職者との接点が不足しています。",
        "求人改善型":"求人内容の魅力が不足している可能性があります。",
        "認知不足型":"会社の認知度が不足しています。",
        "組織改善型":"組織体制の改善余地があります。"
    }

    comment = comments[result_type]


    # 診断日時
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


    data = {
        "診断日時": now,
        "会社名": company,
        "接点": contact,
        "求人力": recruit,
        "認知": recognition,
        "組織力": organization,
        "診断結果": result_type
    }

    # 各回答追加
    data.update(answers)


    file = "diagnosis_results.xlsx"


    if os.path.exists(file):

        df = pd.read_excel(file)
        df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)

    else:

        df = pd.DataFrame([data])


    df.to_excel(file, index=False)


    return render_template(
        "result.html",
        company=company,
        result=result_type,
        comment=comment,
        contact=contact,
        recruit=recruit,
        recognition=recognition,
        organization=organization
    )


app.run(debug=True)