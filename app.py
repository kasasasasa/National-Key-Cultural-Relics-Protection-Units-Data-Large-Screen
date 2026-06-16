#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time : 2020/8/26 14:48
# @Author : way
# @Site :
# @Describe:

from flask import Flask, jsonify, render_template

from data import SourceData


app = Flask(__name__)


@app.route('/')
def index():
    data = SourceData()
    return render_template('index.html', form=data, title=data.title)


@app.route('/api/data')
def api_data():
    data = SourceData()
    return jsonify(data.to_dict())


if __name__ == "__main__":
    app.run(host='127.0.0.1', debug=False)
