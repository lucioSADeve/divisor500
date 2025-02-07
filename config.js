module.exports = {
    session: {
        secret: 'sua-chave-secreta-aqui',
        resave: false,
        saveUninitialized: true,
        cookie: { secure: process.env.NODE_ENV === 'production' }
    }
};