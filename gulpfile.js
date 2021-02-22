"use strict";

var gulp = require("gulp");
var del = require("del");
var typescript = require("gulp-typescript");
var merge = require('merge-stream');
var path = require('path')
var glob = require('glob')

gulp.task('clean', function (cb) {
    return glob('./**/*.ts', function (err, files) {
		var generatedFiles = files.map(function (file) {
		  return file.replace(/.ts$/, '.js*');
		});

		del(generatedFiles, cb);
	});
});

gulp.task('build:InstallMailApp', function () {
    var tsProject = typescript.createProject('./InstallMailApp/tsconfig.json');
    return tsProject.src()
        .pipe(tsProject())
        .pipe(gulp.dest(function(file) {
			return file.base;
		}));
});

gulp.task('build:UninstallMailApp', function () {
    var tsProject = typescript.createProject('./UninstallMailApp/tsconfig.json');
    return tsProject.src()
        .pipe(tsProject())
        .pipe(gulp.dest(function(file) {
			return file.base;
		}));
});

gulp.task('watch:InstallMailApp', gulp.series('build:InstallMailApp'), function () {
    gulp.watch('./InstallMailApp/**/*.ts', gulp.series('build:InstallMailApp'));
});

gulp.task('watch:UninstallMailApp', gulp.series('build:UninstallMailApp'), function () {
    gulp.watch('./UninstallMailApp/**/*.ts', gulp.series('build:UninstallMailApp'));
});

gulp.task('watch', gulp.series('watch:InstallMailApp','watch:UninstallMailApp'));

gulp.task('package:clean', function () {
    return del(gulp.series('package/*/**'));
});

gulp.task('package:copy', gulp.series('package:clean'), function () {
    var main = gulp.src([
                './extension-icon*.png', 'LICENSE', 'README.md', 'vss-extension.json', 'package.json', 
                'docs/**/*', 
				'**/*.xsd', '**/*.wsdl',
                './InstallMailApp/**/*.js', './InstallMailApp/package.json', './InstallMailApp/task.json', './InstallMailApp/icon.png',
                './UninstallMailApp/**/*.js', './UninstallMailApp/package.json', './UninstallMailApp/task.json', './UninstallMailApp/icon.png',
				"!**/node_modules/**/*",],
                 {base: "."})
        .pipe(gulp.dest("./package"));

    return merge(main);	
});

gulp.task('package', gulp.series('package:copy'), function () {	
});

gulp.task('build', gulp.series('build:InstallMailApp','build:UninstallMailApp'));
gulp.task('default', gulp.series('build'));