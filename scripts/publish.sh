#!/bin/bash

set -e

npm config set registry https://registry.npm.taobao.org/

pnpm i --frozen-lockfile

pnpm build

npm config set registry https://registry.npmjs.org/

cd ..

dir

cd ./dist

npm login

npm publish

echo "✅ Publish completed"
