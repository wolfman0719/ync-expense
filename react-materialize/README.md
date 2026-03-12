# reactアプリケーション経費項目編集

CSSフレームワークをMaterializeに変更

## react app テンプレート作成

```% npx create-react-app expense --template typescript```

## materializeインストール

```% cd expense```

```% npm install materialize-css```

## react-router-domインストール

```% npm install react-router-dom```

## axios インストール

```% npm install axios```

## highlight.jsのインストール

```% npm install highlight.js```

## ファイルコピー

以下のファイルをここからダウンロードし、上で作成したテンプレートディレクトリにコピーする
(このディレクトリーに存在しないファイルは、react配下からコピー)

- public

  index.html

  intersystems.css

- src

  index.tsx

  App.tsx

  serverconfig.json

 - components

   ExpenseItem.css

   ExpenseItem.tsx   
   
   Header.tsx

   Query.tsx

   ExpenseItem.tsx

   ExpenseItemList.tsx

  - hooks

    useWindowSize.ts

## serverconfig.jsonの調整

 IRISサーバーのIPアドレス、ポート番号を反映
 (デフォルト　IPアドレス = localhost IPポート番号: 52773)

 ローカルにセットアップした環境では、ポート番号をその環境に合わせて変更する

## reactアプリケーションの起動

- npm start

    Starts the development server.

- npm run build

    Bundles the app into static files for production.

- npm test

    Starts the test runner.

- npm run eject

    Removes this tool and copies build dependencies, configuration files
    and scripts into the app directory. If you do this, you can’t go back!

## CORS設定

開発モード(npm start)で動作させるためには、CORSの設定が必要

### http.confの修正（以下の行を追加）

macOSの場合

```
/opt/homebrew/etc/httpd
```

```
<IfModule mod_headers>
    Header set Access-Control-Allow-Origin "*"
    Header set Access-Control-Allow-Methods "GET,POST,PUT,DELETE,OPTIONS, PATCH"
    Header set Access-Control-Allow-Headers "Content-Type,Authorization,X-Requested-With"
    Header set Access-Control-Allow-Credentials "true"
</IfModule>
```

### IIS

IISの場合は、以下の設定を参考

https://mihono-bourbon.com/iis-cors/

## .htaccessの設定

デプロイの際（npm run build）には.htaccessを作成し、redirectの設定を行う

### http.conf

```
<Directory />
  AllowOverride All
</Directory>
```

### .htaccessの内容

以下の様な内容を記述する

```
RewriteEngine On
RewriteCond %{REQUEST_FILENAME} !-f
RewriteCond %{REQUEST_FILENAME} !-d
RewriteCond %{REQUEST_FILENAME} !-l
RewriteRule ^ index.html [QSA,L]
```

