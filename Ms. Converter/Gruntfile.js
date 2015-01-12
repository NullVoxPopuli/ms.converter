'use strict';

module.exports = function (grunt) {
    // Project configuration.
    grunt.initConfig({
        pkg: grunt.file.readJSON('package.json'),
        compass: {
            dist: {
                options: {
                    cssDir: 'css/',
                    sassDir: 'sass/',
                    outputStyle: 'compressed'
                }
            }
        },
        watch: {
            css: {
                files: '**/*.scss',
                tasks: ['compass']
            }
        }
    });

    // Load the plugins.
    grunt.loadNpmTasks('grunt-contrib-compass');
    grunt.loadNpmTasks('grunt-contrib-watch');

    // default task
    grunt.registerTask('default', ['watch']);
}