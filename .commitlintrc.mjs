export default {
    extends: ['@commitlint/config-conventional'],
    // modify rules, default is below
    rules: {
        'type-enum': [
            2, //RuleConfigSeverity.Error
            'always',
            [
                'feat',
                'fix',
                'docs',
                'style',
                'refactor',
                'test',
                'chore',
                'perf',
                'config'
            ]
        ]
    }
};

// commonJs
// module.exports = {
//     extends: ['@commitlint/config-conventional']
// };

// config-conventional default
// [
//     'build',
//     'chore',
//     'ci',
//     'docs',
//     'feat',
//     'fix',
//     'perf',
//     'refactor',
//     'revert',
//     'style',
//     'test',
// ],