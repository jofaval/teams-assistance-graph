const { readFileSync, writeFileSync } = require('fs');
const { join } = require('path');

// mails generated with: https://www.ipvoid.com/random-email/

function readMails() {
    return readFileSync(join(__dirname, 'mails.txt'), 'utf-8')
        .split(/\r?\n/)
        .map((line) => line.trim())
        .filter(Boolean);
}

function decodeWithEncoding(buffer) {
    if (buffer.length >= 2 && buffer[0] === 0xff && buffer[1] === 0xfe) {
        return {
            content: buffer.slice(2).toString('utf16le'),
            encoding: 'utf16le',
            hasBom: true,
        };
    }

    return {
        content: buffer.toString('utf8'),
        encoding: 'utf8',
        hasBom: false,
    };
}

function encodeWithEncoding(content, encoding, hasBom) {
    const body = Buffer.from(content, encoding);
    if (encoding === 'utf16le' && hasBom) {
        return Buffer.concat([Buffer.from([0xff, 0xfe]), body]);
    }

    if (encoding === 'utf8' && hasBom) {
        return Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), body]);
    }

    return body;
}

function anonymizeEmails(content, replacementMails) {
    const emailRegex = /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi;
    const mapping = new Map();
    let replacementIndex = 0;

    const anonymizedContent = content.replace(emailRegex, (email) => {
        const key = email.toLowerCase();

        if (!mapping.has(key)) {
            if (replacementIndex >= replacementMails.length) {
                throw new Error(
                    `No hay suficientes correos en mails.txt. Necesarios: ${mapping.size + 1}, disponibles: ${replacementMails.length}.`
                );
            }
            mapping.set(key, replacementMails[replacementIndex]);
            replacementIndex += 1;
        }

        return mapping.get(key);
    });

    return {
        anonymizedContent,
        uniqueMapped: mapping.size,
    };
}

function anonymizeParticipantNames(content) {
    const lines = content.split(/\r?\n/);
    const sectionHeaderRegex = /^(\d+)\.\s+/;
    const nameMapping = new Map();
    let currentSection = 0;
    let aliasIndex = 1;

    const anonymizedLines = lines.map((line) => {
        const trimmedLine = line.trim();
        const sectionMatch = trimmedLine.match(sectionHeaderRegex);

        if (sectionMatch) {
            currentSection = Number(sectionMatch[1]);
            return line;
        }

        if (currentSection !== 2 && currentSection !== 3) {
            return line;
        }

        if (trimmedLine === '' || trimmedLine.startsWith('Nombre\t')) {
            return line;
        }

        const columns = line.split('\t');
        if (columns.length < 2 || columns[0].trim() === '') {
            return line;
        }

        const originalName = columns[0].trim();
        const key = originalName.toLowerCase();

        if (!nameMapping.has(key)) {
            const alias = `Participante ${String(aliasIndex).padStart(4, '0')}`;
            nameMapping.set(key, alias);
            aliasIndex += 1;
        }

        columns[0] = nameMapping.get(key);
        return columns.join('\t');
    });

    return {
        anonymizedContent: anonymizedLines.join('\n'),
        uniqueMapped: nameMapping.size,
    };
}

function main() {
    const mails = readMails();
    const inputPath = join(__dirname, 'example-report.csv');
    const outputPath = join(__dirname, 'example-report-anonymized.csv');
    const inputBuffer = readFileSync(inputPath);

    if (inputBuffer.length === 0) {
        throw new Error('example-report.csv está vacío. Añade el contenido del informe y vuelve a ejecutar el script.');
    }

    const { content, encoding, hasBom } = decodeWithEncoding(inputBuffer);
    const { anonymizedContent: nameAnonymizedContent, uniqueMapped: uniqueNamesMapped } = anonymizeParticipantNames(content);
    const { anonymizedContent, uniqueMapped } = anonymizeEmails(nameAnonymizedContent, mails);
    const outputBuffer = encodeWithEncoding(anonymizedContent, encoding, hasBom);

    writeFileSync(outputPath, outputBuffer);

    console.log(`Archivo anonimizado generado en: ${outputPath}`);
    console.log(`Nombres únicos reemplazados: ${uniqueNamesMapped}`);
    console.log(`Correos únicos reemplazados: ${uniqueMapped}`);
}

main();