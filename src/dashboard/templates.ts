export const HTML = `
<div>
    <div class="row">
        <div id="navigation" class="col"></div>
    </div>
    <div id="touRow" class="row">
        <div id="tou" class="col"></div>
    </div>
    <div class="row">
        <div id="filter" class="col mb-4"></div>
    </div>
    <div class="row">
        <div id="table" class="col"></div>
    </div>
</div>`;

export const TOU_HTML = `
<h5>Terms of Service</h5>
<p>In order to use this application you must first agree to the terms of use.</p>
<div class="card" style="width:80%;">
    <div class="card-body" id="touInject"></div>
</div>
<br />
<!-- Agree button will be injected here via script -->`;