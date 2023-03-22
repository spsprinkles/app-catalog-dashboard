'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);
build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
      generatedConfiguration.externals.splice(generatedConfiguration.externals.indexOf('react'), 1);
      generatedConfiguration.externals.splice(generatedConfiguration.externals.indexOf('react-dom'), 1);
      return generatedConfiguration;
  }
});

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.tslintCmd.enabled = false;

build.initialize(require('gulp'));
