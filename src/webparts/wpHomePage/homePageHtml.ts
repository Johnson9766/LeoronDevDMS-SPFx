export default class homeHTML{

public static allElementsHtml:string= ` <div class="main-wrapper w-100 float-start">
    <div class="w-100 float-start leoron-banner-wrapper">
      <div class="w-100 float-start leoron-banner-swiper swiper">
        <div class="swiper-wrapper">
          <!--<div class="leoron-banner-swiper-item swiper-slide">
            <img src="/sites/DevDMS/SiteAssets/resources/images/banner.png"/>
          </div>
          <div class="leoron-banner-swiper-item swiper-slide">
            <img src="/sites/DevDMS/SiteAssets/resources/images/banner.png"/>
          </div>-->
        </div>
         <div class="swiper-pagination"></div>
      </div>
      <div class="leoron-banner-title float-start w-100">
        <div class="container container-leoron mx-auto clearfix px-3 py-4 px-lg-4">
          <p class="text-white text-size-48 font-bold">Digital <br/>Archive</p>
        </div>
      </div>
    </div>
    <div class="w-100 float-start home-page">
      <div class="container-leoron mx-auto clearfix px-3 py-4 px-lg-4">
        <div class="w-100 float-start d-flex flex-column gap-4 home-page-tab-wrapper">
          <div class="home-page-tab-title-wrapper d-flex flex-column flex-sm-row w-100 float-start gap-3">
            <!-- <div data-tab-obj="tabGridData" class="home-page-tab-title flex-fill text-center">Quick Overview</div> 
             <div data-tab-obj="tabGridData" class="home-page-tab-title home-page-tab-title-active flex-fill text-center">Department Functions</div>
             <div data-tab-obj="tabGridData" class="home-page-tab-title flex-fill text-center">Centralized Functions</div>-->
          </div>
          <div class="w-100 float-start home-page-tab-content-wrapper">
           <div class="w-100 float-start home-page-tab-grid-view py-2 gap-3" id="tabGridRoot"></div>
          </div>
        </div>
      </div>
    </div>
  </div>`;

}