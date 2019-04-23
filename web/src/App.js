import React, {Component} from 'react';
import logo from './logo.svg';
import './App.css';
import axios from 'axios';
import {Upload, Button, message, Icon, Modal} from "antd";

const instance = axios.create({
  baseURL: process.env.NODE_ENV === 'development' ? 'http://127.0.0.1:5000' : ''
});

class App extends Component {
  state = {
    filename: '',
    targetUrl: undefined
  };
  error_message = '上传失败，请联系可怜的老公！';
  upload_props = {
    name: 'file',
    action: `${process.env.NODE_ENV === 'development' ? 'http://127.0.0.1:5000' : ''}/server/upload`,
    onChange: ({file}) => {
      if (file.status === 'done') {
        if (file.response.success) {
          this.setState({
            filename: file.response.filename
          });
          message.success('上传成功，请点击处理');
        } else {
          message.error(this.error_message);
        }
      } else if (file.status === 'error') {
        message.error(this.error_message);
      }

    },
  };

  upload = () => {
    instance.post('/server/exe', {filename: this.state.filename}).then((res) => {
      this.setState({
        targetUrl: res.data
      });
      Modal.info({
        title: '处理完成，请点击下方按钮下载',
        content: (
          <div>
            <Button type="primary" icon="cloud-download" onClick={this.downloadUrl}></Button>
          </div>
        )
      });
    }).catch(e => {
      message.error(this.error_message);
    });
  };
  downloadUrl = () => {
    let iframe = document.createElement('iframe');
    iframe.style.display = 'none';
    iframe.src = `${process.env.NODE_ENV === 'development' ? 'http://127.0.0.1:5000' : ''}${this.state.targetUrl}`;
    document.body.appendChild(iframe);
    setTimeout(() => {
      document.body.removeChild(iframe);
    },100);

  };


  render() {
    return (
      <div className="App">
        <header className="App-header">
          <img src={logo} className="App-logo" alt="logo"/>
        </header>
        <div className="content">
          <Upload {...this.upload_props} className="upload">
            <Button>
              <Icon type='upload'/>请点击上传Excel
            </Button>
          </Upload>
          <Button type="primary" disabled={!this.state.filename} onClick={this.upload}>点击处理</Button>
        </div>
      </div>
    );
  }
}

export default App;
