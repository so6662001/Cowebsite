<template>
  <footer class="bg-bg-dark text-white" role="contentinfo">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-16">
      <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-10">
        <div>
          <div class="flex items-center gap-3 mb-6">
            <div class="w-10 h-10 rounded-lg bg-primary flex items-center justify-center">
              <span class="text-white font-bold text-xl">普</span>
            </div>
            <div>
              <h3 class="text-lg font-bold">普讯软件</h3>
              <p class="text-sm text-gray-400">钢铁贸易数字化解决方案</p>
            </div>
          </div>
          <p class="text-gray-400 text-sm leading-relaxed mb-4">
            深耕钢贸行业近20年，累计投入研发近5000万，致力于做全球最受欢迎的钢贸数智赋能平台。
          </p>
          <p class="text-gray-400 text-sm">
            <strong class="text-gray-300">经营理念：</strong>务实·前瞻·安全·有序·开放·适度
          </p>
        </div>

        <div>
          <h4 class="text-base font-semibold mb-5">核心产品</h4>
          <ul class="space-y-3 text-sm text-gray-400">
            <li><router-link to="/products" class="hover:text-white transition-colors">钢贸宝 ERP</router-link></li>
            <li><router-link to="/products" class="hover:text-white transition-colors">钢企通 ERP</router-link></li>
            <li><router-link to="/products" class="hover:text-white transition-colors">库准 生产MES</router-link></li>
            <li><router-link to="/products" class="hover:text-white transition-colors">智能财务系统</router-link></li>
            <li><a href="https://www.huodaizi.com" target="_blank" rel="noopener" class="hover:text-white transition-colors">货袋子赋能平台</a></li>
          </ul>
        </div>

        <div>
          <h4 class="text-base font-semibold mb-5">快速入口</h4>
          <ul class="space-y-3 text-sm text-gray-400">
            <li><router-link to="/solutions" class="hover:text-white transition-colors">解决方案</router-link></li>
            <li><router-link to="/partners" class="hover:text-white transition-colors">合作企业</router-link></li>
            <li><router-link to="/learning" class="hover:text-white transition-colors">学习中心</router-link></li>
            <li><router-link to="/trial" class="hover:text-white transition-colors">申请试用</router-link></li>
            <li><router-link to="/contact" class="hover:text-white transition-colors">联系我们</router-link></li>
          </ul>
        </div>

        <div>
          <h4 class="text-base font-semibold mb-5">联系方式</h4>
          <ul class="space-y-3 text-sm text-gray-400">
            <li class="flex items-start gap-2">
              <svg class="w-4 h-4 mt-0.5 text-primary flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17.657 16.657L13.414 20.9a1.998 1.998 0 01-2.827 0l-4.244-4.243a8 8 0 1111.314 0z"/><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 11a3 3 0 11-6 0 3 3 0 016 0z"/></svg>
              <span>总部：上海<br>研发中心：南京<br>销售中心：南京</span>
            </li>
            <li class="flex items-start gap-2">
              <svg class="w-4 h-4 mt-0.5 text-primary flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 5a2 2 0 012-2h3.28a1 1 0 01.948.684l1.498 4.493a1 1 0 01-.502 1.21l-2.257 1.13a11.042 11.042 0 005.516 5.516l1.13-2.257a1 1 0 011.21-.502l4.493 1.498a1 1 0 01.684.949V19a2 2 0 01-2 2h-1C9.716 21 3 14.284 3 6V5z"/></svg>
              <span>服务中心：025-xxxxxxxx</span>
            </li>
          </ul>
        </div>
      </div>

      <div class="mt-12 pt-8 border-t border-gray-800">
        <div class="flex flex-col md:flex-row justify-between items-center gap-4">
          <p class="text-sm text-gray-500">
            &copy; {{ new Date().getFullYear() }} 南京普讯管理软件有限公司 版权所有
          </p>
          <div class="flex items-center gap-6 text-sm text-gray-500">
            <span>高新技术企业</span>
            <span>5A级信用企业</span>
            <span>100+软件著作权</span>
          </div>
        </div>
      </div>
    </div>
  </footer>
</template>
