<template>
  <div>
    <!-- Hero -->
    <section class="gradient-hero pt-32 pb-20">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 text-center">
        <h1 class="text-4xl lg:text-5xl font-bold text-white mb-6">合作企业</h1>
        <p class="text-xl text-white/80 max-w-3xl mx-auto">
          近3000家钢贸企业的共同选择，让我们用数据说话
        </p>
      </div>
    </section>

    <!-- Breadcrumb -->
    <nav class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4" aria-label="面包屑导航">
      <ol class="flex items-center gap-2 text-sm text-gray-500" itemscope itemtype="https://schema.org/BreadcrumbList">
        <li itemprop="itemListElement" itemscope itemtype="https://schema.org/ListItem">
          <router-link to="/" itemprop="item"><span itemprop="name">首页</span></router-link>
          <meta itemprop="position" content="1" />
        </li>
        <li class="text-gray-300">/</li>
        <li itemprop="itemListElement" itemscope itemtype="https://schema.org/ListItem">
          <span itemprop="name" class="text-gray-900 font-medium">合作企业</span>
          <meta itemprop="position" content="2" />
        </li>
      </ol>
    </nav>

    <!-- Stats -->
    <section class="py-16 bg-white" aria-label="合作数据">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div class="grid grid-cols-2 lg:grid-cols-4 gap-8">
          <div v-for="stat in stats" :key="stat.label"
            class="text-center p-6 rounded-2xl bg-bg-section">
            <div class="text-3xl lg:text-4xl font-bold text-primary mb-2">{{ stat.value }}</div>
            <div class="text-steel-light text-sm">{{ stat.label }}</div>
          </div>
        </div>
      </div>
    </section>

    <!-- Strategic Partners -->
    <section class="py-16 lg:py-24 bg-bg-section" aria-labelledby="strategic-title">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div class="text-center mb-16">
          <h2 id="strategic-title" class="text-3xl font-bold text-gray-900 mb-4">战略合作客户</h2>
          <p class="text-lg text-steel-light max-w-2xl mx-auto">
            以下是与普讯深度合作的战略客户代表，涵盖钢贸、生产制造等多个领域
          </p>
        </div>
        <div class="grid sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6">
          <div v-for="partner in strategicPartners" :key="partner.name"
            class="bg-white rounded-xl p-6 shadow-sm border border-gray-100 card-hover flex flex-col items-center text-center">
            <div class="w-16 h-16 rounded-full bg-primary-light flex items-center justify-center mb-4">
              <span class="text-primary font-bold text-lg">{{ partner.name.charAt(0) }}</span>
            </div>
            <h3 class="font-semibold text-gray-900 mb-1">{{ partner.name }}</h3>
            <span class="text-xs text-steel-light px-3 py-1 bg-gray-100 rounded-full">{{ partner.type }}</span>
          </div>
        </div>
      </div>
    </section>

    <!-- Product Coverage -->
    <section class="py-16 lg:py-24 bg-white" aria-labelledby="coverage-title">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <h2 id="coverage-title" class="text-3xl font-bold text-gray-900 mb-12 text-center">产品合作覆盖</h2>
        <div class="grid md:grid-cols-2 lg:grid-cols-4 gap-8">
          <div v-for="item in coverage" :key="item.product"
            class="bg-bg-section rounded-2xl p-8 text-center">
            <div class="text-4xl mb-4">{{ item.icon }}</div>
            <h3 class="text-lg font-semibold text-gray-900 mb-2">{{ item.product }}</h3>
            <div class="text-3xl font-bold text-primary mb-2">{{ item.count }}</div>
            <p class="text-sm text-steel-light">{{ item.desc }}</p>
          </div>
        </div>
      </div>
    </section>

    <!-- Become Partner -->
    <section class="py-16 lg:py-24 bg-bg-section" aria-labelledby="join-title">
      <div class="max-w-4xl mx-auto px-4 sm:px-6 lg:px-8 text-center">
        <h2 id="join-title" class="text-3xl font-bold text-gray-900 mb-6">期望与我们合作？</h2>
        <p class="text-lg text-steel-light mb-10 max-w-2xl mx-auto">
          无论您是钢贸企业还是钢铁生产制造企业，普讯都能为您提供最适合的数智化解决方案。
          让我们一起携手共创数字化未来。
        </p>
        <div class="flex flex-wrap justify-center gap-4">
          <router-link to="/trial" class="px-8 py-4 bg-primary text-white font-semibold rounded-xl hover:bg-primary-dark transition-colors shadow-lg">
            申请试用
          </router-link>
          <router-link to="/contact" class="px-8 py-4 border-2 border-primary text-primary font-semibold rounded-xl hover:bg-primary-light transition-colors">
            商务合作咨询
          </router-link>
        </div>
      </div>
    </section>
  </div>
</template>

<script setup>
const stats = [
  { value: '3000+', label: '钢贸宝ERP合作企业' },
  { value: '100+', label: '库准合作企业' },
  { value: '10+', label: '生产MES合作企业' },
  { value: '10万吨', label: '货袋子日均库存' }
]

const strategicPartners = [
  { name: '山东乾钢集团', type: '生产制造' },
  { name: '华东钢铁贸易', type: '钢贸企业' },
  { name: '长江钢铁物流', type: '钢贸企业' },
  { name: '南方钢材集团', type: '钢贸企业' },
  { name: '中原钢贸实业', type: '钢贸企业' },
  { name: '东方钢结构公司', type: '生产加工' },
  { name: '北方钢铁贸易', type: '钢贸企业' },
  { name: '西部钢材物流', type: '钢贸企业' }
]

const coverage = [
  { icon: '📊', product: '钢贸宝 ERP', count: '近3000家', desc: '钢铁贸易企业合作使用' },
  { icon: '📦', product: '库准系统', count: '100+家', desc: '标杆企业上线运行' },
  { icon: '🏭', product: '生产MES', count: '10+家', desc: '生产制造企业合作' },
  { icon: '🌐', product: '货袋子平台', count: '10万吨/日', desc: '日均线上库存量' }
]
</script>
