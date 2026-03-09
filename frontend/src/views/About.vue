<template>
  <div>
    <!-- Hero -->
    <section class="gradient-hero pt-32 pb-20" aria-label="关于我们横幅">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 text-center">
        <h1 class="text-4xl lg:text-5xl font-bold text-white mb-6">关于普讯</h1>
        <p class="text-xl text-white/80 max-w-3xl mx-auto">
          深耕钢铁贸易数字化近20年，帮助更多的企业实现更高效率、更有质量的发展
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
          <span itemprop="name" class="text-gray-900 font-medium">关于我们</span>
          <meta itemprop="position" content="2" />
        </li>
      </ol>
    </nav>

    <!-- Company Introduction -->
    <section class="py-16 lg:py-24" aria-labelledby="intro-title">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div class="grid lg:grid-cols-2 gap-16 items-center">
          <div>
            <h2 id="intro-title" class="text-3xl font-bold text-gray-900 mb-6">南京普讯管理软件有限公司</h2>
            <div class="space-y-4 text-steel-light leading-relaxed">
              <p>
                普讯软件是一家<strong class="text-gray-900">钢铁贸易数字化解决方案提供商</strong>和<strong class="text-gray-900">货袋子钢铁交易平台运营商</strong>。
                我们深耕钢贸行业近20年，深入钢贸企业实战运营4年，累计投入研发近5000万，员工近100人，服务团队覆盖全国各地。
              </p>
              <p>
                产品升级以周为单位，快速满足市场和客户需求。服务企业从钢贸企业到钢铁生产制造企业，产品服务从企业经营数智化系统、线上卖货、企业赋能培训到钢神经营俱乐部。
              </p>
              <p>
                <strong class="text-gray-900">我们给企业的不是一套软件，而是一套业务发展的经营系统。</strong>我们服务企业的过程不是简单的培训，而是给你一套经营企业的思路、理念以及落地方法。
              </p>
              <p>
                <strong class="text-gray-900">我们不是一家软件公司，而是一个赋能平台</strong>，让更多企业享受更好的服务。
              </p>
            </div>
          </div>
          <div class="bg-gradient-to-br from-primary-light to-blue-50 rounded-2xl p-8">
            <div class="grid grid-cols-2 gap-6">
              <div v-for="stat in stats" :key="stat.label" class="bg-white rounded-xl p-6 text-center shadow-sm">
                <div class="text-2xl lg:text-3xl font-bold text-primary mb-1">{{ stat.value }}</div>
                <div class="text-sm text-steel-light">{{ stat.label }}</div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </section>

    <!-- Philosophy -->
    <section class="py-16 lg:py-24 bg-bg-section" aria-labelledby="philosophy-title">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <h2 id="philosophy-title" class="text-3xl font-bold text-gray-900 mb-12 text-center">企业理念</h2>
        <div class="grid md:grid-cols-2 lg:grid-cols-4 gap-8">
          <div v-for="p in philosophies" :key="p.q" class="bg-white rounded-2xl p-6 shadow-sm border border-gray-100">
            <h3 class="text-primary font-semibold mb-3">{{ p.q }}</h3>
            <p class="text-steel-light text-sm leading-relaxed">{{ p.a }}</p>
          </div>
        </div>
      </div>
    </section>

    <!-- Timeline -->
    <section class="py-16 lg:py-24 bg-white" aria-labelledby="timeline-title">
      <div class="max-w-4xl mx-auto px-4 sm:px-6 lg:px-8">
        <h2 id="timeline-title" class="text-3xl font-bold text-gray-900 mb-12 text-center">发展历程</h2>
        <div class="relative">
          <div class="absolute left-4 md:left-1/2 top-0 bottom-0 w-0.5 bg-gray-200 -translate-x-1/2"></div>
          <div v-for="(event, i) in timeline" :key="event.year" class="relative mb-12 last:mb-0">
            <div class="flex items-start gap-8" :class="i % 2 === 0 ? 'md:flex-row' : 'md:flex-row-reverse'">
              <div class="hidden md:block w-1/2" :class="i % 2 === 0 ? 'text-right pr-12' : 'text-left pl-12'">
                <div class="inline-block bg-white rounded-xl p-6 shadow-sm border border-gray-100">
                  <span class="text-2xl font-bold text-primary">{{ event.year }}</span>
                  <p class="text-steel-light text-sm mt-2">{{ event.desc }}</p>
                </div>
              </div>
              <div class="absolute left-4 md:left-1/2 w-4 h-4 bg-primary rounded-full -translate-x-1/2 mt-2 ring-4 ring-white"></div>
              <div class="md:hidden pl-10">
                <div class="bg-white rounded-xl p-5 shadow-sm border border-gray-100">
                  <span class="text-xl font-bold text-primary">{{ event.year }}</span>
                  <p class="text-steel-light text-sm mt-2">{{ event.desc }}</p>
                </div>
              </div>
              <div class="hidden md:block w-1/2"></div>
            </div>
          </div>
        </div>
      </div>
    </section>

    <!-- Honors -->
    <section class="py-16 lg:py-24 bg-bg-section" aria-labelledby="honors-title">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <h2 id="honors-title" class="text-3xl font-bold text-gray-900 mb-12 text-center">荣誉资质</h2>
        <div class="grid sm:grid-cols-2 lg:grid-cols-4 gap-8">
          <div v-for="honor in honors" :key="honor.title"
            class="bg-white rounded-2xl p-8 text-center shadow-sm border border-gray-100 card-hover">
            <div class="w-16 h-16 mx-auto mb-4 rounded-xl flex items-center justify-center text-3xl" :class="honor.bg">
              {{ honor.icon }}
            </div>
            <h3 class="text-lg font-semibold text-gray-900 mb-2">{{ honor.title }}</h3>
            <p class="text-sm text-steel-light">{{ honor.desc }}</p>
          </div>
        </div>
      </div>
    </section>

    <!-- CTA -->
    <section class="py-16 gradient-hero text-center">
      <div class="max-w-4xl mx-auto px-4">
        <h2 class="text-3xl font-bold text-white mb-6">与普讯同行，共创数字化未来</h2>
        <p class="text-white/80 mb-8">帮助客户发展，帮助伙伴发展，让与我们共同合作的人赚到钱</p>
        <router-link to="/trial" class="inline-block px-8 py-4 bg-white text-primary font-semibold rounded-xl hover:bg-gray-100 transition-all shadow-lg">
          立即开始合作
        </router-link>
      </div>
    </section>
  </div>
</template>

<script setup>
const stats = [
  { value: '近20年', label: '深耕行业' },
  { value: '3000+', label: '合作企业' },
  { value: '5000万', label: '研发投入' },
  { value: '100人', label: '团队规模' },
  { value: '10000+', label: '日活用户' },
  { value: '100+', label: '软件著作权' },
  { value: '10万吨', label: '日均线上库存' },
  { value: '全国', label: '服务覆盖' }
]

const philosophies = [
  { q: '普讯秉承什么理念？', a: '务实、前瞻、安全、有序、开放、适度的经营理念。' },
  { q: '普讯为什么而存在？', a: '帮助更多的企业实现更高效率、更有质量的发展；帮助更多的伙伴实现人生价值；让与我们一起共同合作的人赚到钱。' },
  { q: '普讯要往哪里去？', a: '做全球最受欢迎的钢贸数智赋能平台。' },
  { q: '普讯的经营理念？', a: '帮助客户发展，帮助伙伴发展。' }
]

const timeline = [
  { year: '2010', desc: '南京普讯管理软件有限公司正式成立，开启钢铁贸易数字化之路。' },
  { year: '2016', desc: '完成全国50多家大型钢贸企业客户的合作，深入行业实战运营。' },
  { year: '2023', desc: '完成全国2000+家钢贸企业的钢贸宝ERP合作，并上线山东乾钢生产制造企业MES数字化系统。' },
  { year: '2025', desc: '完成100多家库准标杆企业上线，货袋子平台在标杆企业完成试运行。' }
]

const honors = [
  { icon: '🏆', title: '高新技术企业', desc: '荣获2家高新企业认证', bg: 'bg-amber-100' },
  { icon: '⭐', title: '5A级信用企业', desc: '企业信用等级5A认证', bg: 'bg-blue-100' },
  { icon: '📜', title: '100+著作权', desc: '拥有100多个产品著作权', bg: 'bg-green-100' },
  { icon: '🔬', title: '4项专利', desc: '4项专利（在审中）', bg: 'bg-purple-100' }
]
</script>
