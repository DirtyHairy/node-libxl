import { readFileSync } from 'fs';

let code = readFileSync(process.argv[2], 'utf8');
let tokens = code.match(/[\w]+|[{}]/gm);

let state = 'scan';

let enums = new Map();

let name = '';
for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];

    switch (state) {
        case 'scan':
            if (token === 'enum') state = 'enum';
            break;

        case 'enum':
            if (/\w+/.test(token)) {
                name = token;

                enums.set(name, []);
                state = 'name';
            }

            break;

        case 'name':
            if (token === '{') {
                state = 'values';
                break;
            }

            throw new Error(`unexpected token after enum name ${token}`);

        case 'values':
            if (token === '}') {
                state = 'scan';
                break;
            }

            if (/\w+/.test(token)) {
                if (!/^\d/.test(token)) enums.get(name).push(token);
                break;
            }

            throw new Error(`unexpected token in enum values ${token}`);
    }
}

for (let [, values] of enums) {
    console.log(`// ${name}`);
    for (let value of values) {
        console.log(`NODE_DEFINE_CONSTANT(exports, ${value});`);
    }

    console.log();
}
